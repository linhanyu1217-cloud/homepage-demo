#Requires -Version 5.1
<#
================================================================================
 run_weekly.ps1 — 地产周报 Agent 每周运行脚本
================================================================================
 版本：v1.0 ／ 建立日期：2026-09-16

 作用
   1. 自动进入工作区 C:\Users\15258\Desktop\地产周报Agent
   2. 从项目根目录 .env 加载 ANTHROPIC_BASE_URL / ANTHROPIC_AUTH_TOKEN /
      ANTHROPIC_MODEL
   3. 全程不输出 AUTH_TOKEN（日志再做一次脱敏兜底）
   4. 读取 提示词\地产周报任务卡.txt
   5. 用 `claude -p` 非交互运行 Agent
   6. 自动建立 05_输出\YYYY-Wxx
   7. 日志写入 日志\
   8. 出错记录完整错误信息，但不记录 API Key
   9. 失败后最多自动重试一次；第二次失败即停止（不无限重试）

 用法
   .\脚本\run_weekly.ps1                    # 正常运行
   .\脚本\run_weekly.ps1 -DryRun            # 只做校验，不调用模型
   .\脚本\run_weekly.ps1 -Week 2026-W37     # 指定周次标签
   .\脚本\run_weekly.ps1 -EndDate 2026-09-11# 指定窗口末日（周五）

 安全边界
   - 只写本工作区；历史「实习」目录只读
   - 不注册计划任务、不发邮件、不上传
================================================================================
#>
[CmdletBinding()]
param(
    # 周次标签，形如 2026-W38。不传则由窗口末日推算。
    [ValidatePattern('^\d{4}-W\d{2}$')]
    [string]$Week,

    # 报告窗口末日（应为周五）。不传则自动取「最近的、已结束的」周五。
    [datetime]$EndDate,

    # 任务卡路径覆盖
    [string]$TaskCard,

    # 只做前置校验与目录准备，不调用 claude
    [switch]$DryRun,

    # 不把历史「实习」目录加入可访问范围
    [switch]$NoHistoryDir,

    # 传给 claude 的工具白名单
    [string]$AllowedTools = 'Read,Write,Edit,Glob,Grep,Bash(python *)'
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ 常量
$BASE        = 'C:\Users\15258\Desktop\地产周报Agent'
$HISTORY_DIR = 'C:\Users\15258\Desktop\实习'
$MAX_ATTEMPTS = 2          # 失败后最多重试一次 → 总共 2 次

# ------------------------------------------------------------------ 编码
# PS 5.1 默认 $OutputEncoding 是 ASCII，会把中文管道内容毁掉，必须显式设成 UTF-8。
$OutputEncoding = New-Object System.Text.UTF8Encoding $false
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

# ------------------------------------------------------------------ 日志
$script:LogPath   = $null
$script:LogDir    = Join-Path $BASE '日志'
$script:Secrets   = @()      # 需要脱敏的字符串（仅存在于内存，不落盘）

function Protect-LogText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    foreach ($s in $script:Secrets) {
        if ($s -and $s.Length -ge 8) { $Text = $Text.Replace($s, '****REDACTED****') }
    }
    return $Text
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','STEP')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, (Protect-LogText $Message)
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        default { Write-Host $line }
    }
    if ($script:LogPath) {
        try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 } catch { }
    }
}

# ------------------------------------------------------------------ ISO 周
# Windows PowerShell 5.1 没有 [System.Globalization.ISOWeek]，自行实现。
# 算法：取该日期所在 ISO 周的星期四，用它定 ISO 年，再数周数。
# 已对 2026-09-11→W37 / 2027-01-01→2026-W53 / 2025-12-29→2026-W01 等
# 边界用例验证通过。
function Get-IsoWeek {
    param([Parameter(Mandatory)][datetime]$Date)
    $d = $Date.Date
    $isoDow = [int]$d.DayOfWeek
    if ($isoDow -eq 0) { $isoDow = 7 }          # 周日 → 7
    $thursday = $d.AddDays(4 - $isoDow)
    $isoYear  = $thursday.Year
    $jan4     = Get-Date -Year $isoYear -Month 1 -Day 4
    $j4dow    = [int]$jan4.DayOfWeek
    if ($j4dow -eq 0) { $j4dow = 7 }
    $week1Mon = $jan4.AddDays(-($j4dow - 1))
    $week     = [int][math]::Floor(($thursday - $week1Mon).TotalDays / 7) + 1
    return [pscustomobject]@{ IsoYear = $isoYear; IsoWeek = $week }
}

# ------------------------------------------------------------------ .env
function Read-DotEnv {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw ".env 不存在：$Path" }
    $vars = @{}
    foreach ($raw in [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)) {
        $line = $raw.Trim().TrimStart([char]0xFEFF)      # 去 BOM
        if ($line -eq '' -or $line.StartsWith('#')) { continue }
        if ($line.StartsWith('export ')) { $line = $line.Substring(7).Trim() }
        $i = $line.IndexOf('=')
        if ($i -lt 1) { continue }
        $k = $line.Substring(0, $i).Trim()
        $v = $line.Substring($i + 1).Trim()
        if ($v.Length -ge 2) {
            $pair = ($v[0] -eq '"' -and $v[-1] -eq '"') -or ($v[0] -eq "'" -and $v[-1] -eq "'")
            if ($pair) { $v = $v.Substring(1, $v.Length - 2) }
        }
        if ($k) { $vars[$k] = $v }
    }
    return $vars
}

function Protect-Secret {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '(未设置)' }
    if ($Value.Length -le 6) { return '****' }
    return $Value.Substring(0, 6) + '****'
}

# ------------------------------------------------------------------ claude 定位
# Get-Command 在 Windows 上常常只解析到 npm 的 claude.ps1 包装脚本。
# .ps1 包装脚本转发 stdin 时行为不够可控，优先使用原生 claude.exe。
function Resolve-ClaudeExe {
    $cmd = Get-Command claude -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }
    $src = $cmd.Source
    if ([string]::IsNullOrEmpty($src)) { return $null }

    # 已经是原生可执行文件
    if ($src -notmatch '\.(ps1|cmd|bat)$') { return $src }

    # 从 npm 包装脚本推导真正的可执行文件
    $npmDir = Split-Path -Parent $src
    $candidates = @(
        (Join-Path $npmDir 'node_modules\@anthropic-ai\claude-code\bin\claude.exe')
        (Join-Path $npmDir 'claude.exe')
        (Join-Path $npmDir 'claude.cmd')
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return $src   # 兜底：仍用包装脚本（其本身支持管道输入）
}

# ------------------------------------------------------------------ Agent 状态窗口
# 正式开工时自动弹出一个**独立**的状态窗口（`脚本\查看Agent状态.ps1`）。
#   · 用 Start-Process 起一个**自己的 PowerShell 进程**：不阻塞本脚本，
#     也不影响 claude 主任务，主任务照常往下走；
#   · 窗口开在自己的进程里，Agent 结束后**不强制关闭**，由使用者自己关；
#   · 重复运行不会弹出第二个窗口（判定见 Test-StatusWindowOpen）；
#   · `-DryRun` 与「任务卡为空」等提前退出的路径**都不会**走到这里。
$STATUS_SCRIPT = Join-Path $BASE '脚本\查看Agent状态.ps1'
$STATUS_LOCK   = Join-Path $BASE '04_工作区\locks\agent_status_window.lock'
# 必须与 查看Agent状态.ps1 里 $form.Text 完全一致，否则按标题找不到窗口
$STATUS_TITLE  = '地产周报 Agent 状态'

function Get-StatusWindowProcess {
    # 直接找「已经开着的状态窗口」：按窗口标题认，比 PID 可靠
    # （PID 会在一周内被系统回收，锁文件里的旧 PID 可能指向别的进程）
    return @(Get-Process -ErrorAction SilentlyContinue |
             Where-Object { $_.MainWindowTitle -eq $STATUS_TITLE })
}

function Read-StatusLockPid {
    # 锁文件里的 PID：仅用于诊断与测试收尾，不作为「是否已开窗」的主判据
    if (-not (Test-Path -LiteralPath $STATUS_LOCK)) { return 0 }
    try { $raw = (Get-Content -LiteralPath $STATUS_LOCK -Raw -ErrorAction Stop).Trim() } catch { return 0 }
    $n = 0
    if ([int]::TryParse($raw, [ref]$n)) { return $n }
    return 0
}

function Test-StatusWindowOpen {
    # ① 屏幕上是否已有这个标题的窗口（主判据）
    if ((Get-StatusWindowProcess).Count -gt 0) { return $true }
    # ② 兜底：窗口可能还没建好，但锁文件里的 PID 仍然活着
    $lockPid = Read-StatusLockPid
    if ($lockPid -gt 0 -and (Get-Process -Id $lockPid -ErrorAction SilentlyContinue)) { return $true }
    return $false
}

function Show-AgentStatusWindow {
    if (-not (Test-Path -LiteralPath $STATUS_SCRIPT)) {
        Write-Log -Level WARN -Message "找不到状态窗口脚本，跳过弹窗（不影响主任务）：$STATUS_SCRIPT"
        return
    }

    $existing = Get-StatusWindowProcess
    if ($existing.Count -gt 0) {
        Write-Log -Message ("状态窗口已在运行（PID {0}），不重复弹出。" -f ($existing[0].Id))
        return
    }
    if (Test-StatusWindowOpen) {
        Write-Log -Message '状态窗口已在运行（锁文件判据），不重复弹出。'
        return
    }

    # 确认没有窗口和活进程后，锁文件必然是过期残留，先清掉
    if (Test-Path -LiteralPath $STATUS_LOCK) {
        Remove-Item -LiteralPath $STATUS_LOCK -Force -ErrorAction SilentlyContinue
    }
    $lockDir = Split-Path -Parent $STATUS_LOCK
    if (-not (Test-Path -LiteralPath $lockDir)) {
        New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
    }

    try {
        $proc = Start-Process -FilePath 'powershell.exe' `
            -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $STATUS_SCRIPT)) `
            -WorkingDirectory $BASE -PassThru
        Set-Content -LiteralPath $STATUS_LOCK -Value $proc.Id -Encoding ASCII
        Write-Log -Level OK -Message ("已弹出 Agent 状态窗口（PID {0}，独立进程，不阻塞主任务）。" -f $proc.Id)
    } catch {
        # 弹窗失败绝不能影响主任务
        Write-Log -Level WARN -Message ("状态窗口启动失败（不影响主任务）：{0}" -f $_.Exception.Message)
    }
}

# ================================================================== 主流程
$startTime = Get-Date
Write-Host ''
Write-Host '==================== 地产周报 Agent · 每周运行 ====================' -ForegroundColor Cyan
Write-Host ''

# --- 1. 进入工作区 -------------------------------------------------
Set-Location -LiteralPath $BASE
Write-Host ("工作区：{0}" -f $BASE)

# --- 2. 日志文件 ---------------------------------------------------
if (-not (Test-Path -LiteralPath $script:LogDir)) {
    New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
}
$stamp = $startTime.ToString('yyyyMMdd-HHmmss')
$script:LogPath = Join-Path $script:LogDir ("run_$stamp.log")
Write-Log -Level STEP -Message "===== 地产周报 Agent 运行开始 ====="
Write-Log -Message "工作区：$BASE"
Write-Log -Message "日志文件：$($script:LogPath)"

# --- 3. 读取 .env --------------------------------------------------
$envPath = Join-Path $BASE '.env'
Write-Log -Level STEP -Message "[1/6] 读取 .env"
try {
    $envVars = Read-DotEnv -Path $envPath
} catch {
    Write-Log -Level ERROR -Message "读取 .env 失败：$($_.Exception.Message)"
    exit 2
}

$required = @('ANTHROPIC_BASE_URL', 'ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_MODEL')
$missing  = @($required | Where-Object { -not $envVars.ContainsKey($_) -or [string]::IsNullOrWhiteSpace($envVars[$_]) })
if ($missing.Count -gt 0) {
    Write-Log -Level ERROR -Message "缺少必需变量：$($missing -join ', ')"
    exit 2
}

# 把密钥放进内存脱敏表 —— 之后任何写日志的文本都会被替换掉
$script:Secrets = @($envVars['ANTHROPIC_AUTH_TOKEN'])

# 注入到当前进程环境，子进程 claude 会继承
$env:ANTHROPIC_BASE_URL   = $envVars['ANTHROPIC_BASE_URL']
$env:ANTHROPIC_AUTH_TOKEN = $envVars['ANTHROPIC_AUTH_TOKEN']
$env:ANTHROPIC_MODEL      = $envVars['ANTHROPIC_MODEL']

Write-Log -Level OK  -Message "ANTHROPIC_BASE_URL = $($envVars['ANTHROPIC_BASE_URL'])"
Write-Log -Level OK  -Message "ANTHROPIC_AUTH_TOKEN = $(Protect-Secret $envVars['ANTHROPIC_AUTH_TOKEN'])  （已遮蔽，永不完整输出）"
Write-Log -Level OK  -Message "ANTHROPIC_MODEL = $($envVars['ANTHROPIC_MODEL'])"

# --- 4. 周次与输出目录 ---------------------------------------------
Write-Log -Level STEP -Message "[2/6] 计算周次与输出目录"

if (-not $PSBoundParameters.ContainsKey('EndDate')) {
    # 报告窗口 = 「周六 → 周五」。取最近的、已结束的周五：
    #   今天是周五 → 用今天；否则往前回退到最近的周五
    $today = (Get-Date).Date
    $dow   = [int]$today.DayOfWeek                 # 周五 = 5
    $back  = ($dow - 5 + 7) % 7
    $EndDate = if ($back -eq 0) { $today } else { $today.AddDays(-$back) }
}
$EndDate = $EndDate.Date
$StartDate = $EndDate.AddDays(-6)

$iso = Get-IsoWeek -Date $EndDate
if (-not $Week) { $Week = '{0}-W{1:d2}' -f $iso.IsoYear, $iso.IsoWeek }

$reportDir = Join-Path (Join-Path $BASE '05_输出') $Week
Write-Log -Message "报告窗口：$($StartDate.ToString('yyyy-MM-dd'))（周六） → $($EndDate.ToString('yyyy-MM-dd'))（周五）"
Write-Log -Message "周次标签：$Week"

if (-not (Test-Path -LiteralPath $reportDir)) {
    New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
    Write-Log -Level OK -Message "已建立输出目录：$reportDir"
} else {
    Write-Log -Message "输出目录已存在：$reportDir"
}

$env:REPORT_WEEK  = $Week
$env:REPORT_DIR   = $reportDir
$env:REPORT_START = $StartDate.ToString('yyyy-MM-dd')
$env:REPORT_END   = $EndDate.ToString('yyyy-MM-dd')

# --- 4b. 本周输入目录（按周次子文件夹）-------------------------------
# 规则：周次标签 2026-W37  ->  W37公司公告
#   02_待处理\A股公告\W37公司公告
#   02_待处理\港股公告\W37公司公告
# 只读本周目录；绝不扫描其他周次目录；目录缺失**不报错、不退出**，
# 仅记为「本周该市场无输入」，由 Agent 写入 run_log.md。

$mTag = [regex]::Match($Week, 'W(\d{1,2})\s*$')
if (-not $mTag.Success) {
    Write-Log -Level ERROR -Message "周次标签无法解析出 Wnn，无法定位输入目录：$Week（期望形如 2026-W37）"
    exit 5
}
$weekTag   = 'W{0:d2}' -f [int]$mTag.Groups[1].Value
$weekSub   = "${weekTag}公司公告"

$inAShare = Join-Path $BASE "02_待处理\A股公告\$weekSub"
$inHK     = Join-Path $BASE "02_待处理\港股公告\$weekSub"

Write-Log -Message "本周输入子目录：$weekSub（由周次标签 $Week 推导）"

$inputDirs = [ordered]@{
    'A股公告' = $inAShare
    '港股公告' = $inHK
}
$inputMissing = @()
$marketFiles  = @{}
foreach ($mkt in $inputDirs.Keys) {
    $dir = $inputDirs[$mkt]
    $rel = $dir.Substring($BASE.Length).TrimStart('\')
    if (Test-Path -LiteralPath $dir -PathType Container) {
        $files = @(Get-ChildItem -LiteralPath $dir -File -Recurse -ErrorAction SilentlyContinue)
        $marketFiles[$mkt] = $files
        if ($files.Count -gt 0) {
            Write-Log -Level OK -Message "  ✔ $rel  （原始 $($files.Count) 个文件）"
        } else {
            Write-Log -Level WARN -Message "  ○ $rel 存在但为空 → 记为「本周该市场无输入」"
            $inputMissing += "$mkt(空)"
        }
    } else {
        $marketFiles[$mkt] = @()
        Write-Log -Level WARN -Message "  ○ $rel 不存在 → 记为「本周该市场无输入」（不报错，不退出）"
        $inputMissing += "$mkt(缺目录)"
    }
}

# 显式声明被忽略的其他周次目录，证明未把上周/下周公告混入本周
foreach ($mkt in @('A股公告', '港股公告')) {
    $parent = Join-Path $BASE "02_待处理\$mkt"
    if (Test-Path -LiteralPath $parent -PathType Container) {
        $others = @(Get-ChildItem -LiteralPath $parent -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -ne $weekSub })
        if ($others.Count -gt 0) {
            $names = ($others | ForEach-Object { $_.Name }) -join '、'
            Write-Log -Message "  · 已忽略其他周次目录（$mkt）：$names"
        }
    }
}

if ($inputMissing.Count -gt 0) {
    Write-Log -Level WARN -Message "本周无输入的市场：$($inputMissing -join '、')（非阻塞，流程照常执行）"
}

# --- 4c. 跨市场重复检测（只标记忽略，绝不删除 / 移动 / 改名）---------
# 规则：以 **A股目录为基准**，港股目录中「同名 + 同大小 + 同哈希」的文件
#       标记为【忽略：跨市场重复文件】，**不计入港股有效文件数**；
#       A股一侧保留原文件，数量不变。
# 判定顺序按使用者要求：先比文件名 → 再比文件大小 → 可行时再比哈希。
# 只影响「处理阶段」，磁盘上的文件一个都不动。

$dupList  = New-Object System.Collections.ArrayList
$nameOnly = New-Object System.Collections.ArrayList
$aFiles   = @($marketFiles['A股公告'])
$hkFiles  = @($marketFiles['港股公告'])

if ($aFiles.Count -gt 0 -and $hkFiles.Count -gt 0) {
    $aIndex = @{}
    foreach ($f in $aFiles) {
        if (-not $aIndex.ContainsKey($f.Name)) { $aIndex[$f.Name] = $f }
    }
    foreach ($h in $hkFiles) {
        if (-not $aIndex.ContainsKey($h.Name)) { continue }      # 1) 文件名
        $a = $aIndex[$h.Name]
        if ($a.Length -ne $h.Length) {                            # 2) 文件大小
            [void]$nameOnly.Add("$($h.Name)（同名但大小不同：A股 $($a.Length) B / 港股 $($h.Length) B）")
            continue
        }
        $shaA = $null; $shaH = $null                              # 3) 哈希（可行时）
        try { $shaA = (Get-FileHash -LiteralPath $a.FullName -Algorithm SHA256).Hash } catch { }
        try { $shaH = (Get-FileHash -LiteralPath $h.FullName -Algorithm SHA256).Hash } catch { }
        if ($shaA -and $shaH) {
            if ($shaA -ne $shaH) {
                [void]$nameOnly.Add("$($h.Name)（同名同大小但内容不同）")
                continue
            }
            $basis = '同名+同大小+同哈希'
        } else {
            $basis = '同名+同大小（哈希不可用，未比对内容）'
        }
        [void]$dupList.Add([pscustomobject]@{
            '忽略文件路径' = $h.FullName.Substring($BASE.Length).TrimStart('\')
            '对应A股文件'  = $a.FullName.Substring($BASE.Length).TrimStart('\')
            '文件名'       = $h.Name
            '字节数'       = $h.Length
            'sha256'       = $(if ($shaH) { $shaH } else { '' })
            '判定依据'     = $basis
            '处理方式'     = '忽略：跨市场重复文件（不删除、不移动、不改名）'
        })
    }
}

# 清单落盘，供 run_log.md 引用
$dupCsv = Join-Path $reportDir '跨市场重复清单.csv'
if ($dupList.Count -gt 0) {
    $dupList | Export-Csv -LiteralPath $dupCsv -NoTypeInformation -Encoding UTF8
}

$aEffective  = $aFiles.Count
$hkEffective = $hkFiles.Count - $dupList.Count
if ($hkEffective -lt 0) { $hkEffective = 0 }

if ($dupList.Count -gt 0) {
    Write-Log -Level WARN -Message "跨市场重复：港股目录中 $($dupList.Count) 个文件与 A股目录完全重复 → 【忽略：跨市场重复文件】"
    Write-Log -Level WARN -Message "  仅在处理阶段忽略，磁盘文件不删除、不移动；清单：$dupCsv"
} elseif ($aFiles.Count -gt 0 -and $hkFiles.Count -gt 0) {
    Write-Log -Level OK -Message "跨市场重复：未发现重复文件"
}
if ($nameOnly.Count -gt 0) {
    Write-Log -Level WARN -Message "同名但未判定为完全重复（照常处理）：$($nameOnly.Count) 个"
    foreach ($x in $nameOnly) { Write-Log -Message "    · $x" }
}
Write-Log -Level OK -Message "本周有效输入：A股 $aEffective 个 / 港股 $hkEffective 个 / 忽略重复 $($dupList.Count) 个"

# 注入给 claude 子进程
$env:REPORT_WEEK_TAG = $weekTag
$env:REPORT_IN_A     = $inAShare
$env:REPORT_IN_HK    = $inHK
$env:REPORT_A_EFFECTIVE  = "$aEffective"
$env:REPORT_HK_EFFECTIVE = "$hkEffective"
$env:REPORT_DUP_COUNT    = "$($dupList.Count)"
$env:REPORT_DUP_CSV      = $(if ($dupList.Count -gt 0) { $dupCsv } else { '' })

# 新闻资讯 / 研报 保持平铺目录，允许为空（非阻塞待扩展）
$env:REPORT_IN_NEWS  = Join-Path $BASE '02_待处理\新闻资讯'
$env:REPORT_IN_REPORT = Join-Path $BASE '02_待处理\研报'

# --- 5. 前置校验 ---------------------------------------------------
Write-Log -Level STEP -Message "[3/6] 前置校验"

$ruleFiles = @(
    '00_规则\工作地图.md'
    '00_规则\Agent规则.md'
    '00_规则\行业公司池规则.md'
    '00_规则\重点关注池.md'
    '04_工作区\行业公司池.csv'
    '04_工作区\历史覆盖统计.csv'
    '01_模板\公司公告周报模板.docx'
)
$missingRules = @()
foreach ($rel in $ruleFiles) {
    $full = Join-Path $BASE $rel
    if (Test-Path -LiteralPath $full) {
        Write-Log -Level OK -Message "  ✔ $rel"
    } else {
        Write-Log -Level WARN -Message "  ✘ 缺失：$rel  （将记入日志，不伪造内容）"
        $missingRules += $rel
    }
}
if ($missingRules.Count -gt 0) {
    Write-Log -Level WARN -Message "有 $($missingRules.Count) 个规则文件缺失，Agent 会在 run_log.md 中如实记录。"
}

if (-not $TaskCard) { $TaskCard = Join-Path $BASE '提示词\地产周报任务卡.txt' }
if (-not (Test-Path -LiteralPath $TaskCard)) {
    Write-Log -Level ERROR -Message "任务卡不存在：$TaskCard"
    exit 3
}
Write-Log -Level OK -Message "任务卡：$TaskCard"

# .env 是否已被 gitignore
$gitignore = Join-Path $BASE '.gitignore'
if (Test-Path -LiteralPath $gitignore) {
    $gi = Get-Content -LiteralPath $gitignore -Raw -Encoding UTF8
    if ($gi -match '(?m)^\s*\.env\s*$') {
        Write-Log -Level OK -Message ".gitignore 已忽略 .env"
    } else {
        Write-Log -Level WARN -Message ".gitignore 未发现独立的 .env 规则，请检查"
    }
} else {
    Write-Log -Level WARN -Message "未找到 .gitignore"
}

# claude CLI
$claudeExe = Resolve-ClaudeExe
if (-not $claudeExe) {
    Write-Log -Level ERROR -Message "未找到 claude 命令，请确认 Claude Code CLI 已安装并在 PATH 中"
    exit 4
}
Write-Log -Level OK -Message "claude 可执行文件：$claudeExe"

# Python 依赖检查（只报告，绝不自动安装）
$depScript = Join-Path $BASE '脚本\检查依赖.py'
if (Test-Path -LiteralPath $depScript) {
    try {
        $depOut = & python $depScript 2>&1
        $depMissing = 0
        foreach ($ln in $depOut) {
            $parts = "$ln".Split('|')
            if ($parts.Count -lt 2) { continue }
            switch ($parts[0]) {
                'OK'   { Write-Log -Level OK   -Message "  依赖 ✔ $($parts[1]) $($parts[2])" }
                'MISS' {
                    Write-Log -Level WARN -Message "  依赖 ✘ $($parts[1]) 未安装（不得自行安装，需人工确认）"
                    $depMissing++
                }
                default { }
            }
        }
        if ($depMissing -gt 0) {
            Write-Log -Level WARN -Message "有 $depMissing 个依赖缺失，相关产出会被跳过。"
        }
    } catch {
        Write-Log -Level WARN -Message "Python 依赖检查失败：$($_.Exception.Message)"
    }
} else {
    Write-Log -Level WARN -Message "未找到 脚本\检查依赖.py，跳过依赖检查"
}

# 输入投递目录（只确保**顶层**目录存在）
# 注意：**不自动创建周次子目录**（如 W38公司公告）—— 由使用者按周次投放；
#       缺失时按「本周该市场无输入」处理，不报错、不退出。
foreach ($d in @('A股公告','港股公告','新闻资讯','研报')) {
    $p = Join-Path $BASE "02_待处理\$d"
    if (-not (Test-Path -LiteralPath $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
        Write-Log -Message "  已创建投递目录：02_待处理\$d"
    }
}

# --- 6. 调用 claude ------------------------------------------------
Write-Log -Level STEP -Message "[4/6] 组装 claude 调用参数"

$claudeArgs = @(
    '-p'
    '--output-format', 'text'
    '--permission-mode', 'acceptEdits'
    '--allowedTools', $AllowedTools
    '--model', $env:ANTHROPIC_MODEL
)
if (-not $NoHistoryDir -and (Test-Path -LiteralPath $HISTORY_DIR)) {
    $claudeArgs += @('--add-dir', $HISTORY_DIR)
    Write-Log -Message "已加入只读历史资料目录：$HISTORY_DIR（规则要求只读，写入被禁止）"
}

Write-Log -Message "claude 参数：$($claudeArgs -join ' ')"
Write-Log -Message "工具白名单：$AllowedTools"

if ($DryRun) {
    Write-Log -Level WARN -Message "-DryRun 已指定：完成全部校验与目录准备，但不调用模型。"
    Write-Log -Level OK   -Message "===== DryRun 结束（未调用 claude）====="
    Write-Host ''
    Write-Host "DryRun 通过。输出目录：$reportDir" -ForegroundColor Green
    Write-Host "日志：$($script:LogPath)" -ForegroundColor Green
    exit 0
}

$prompt = Get-Content -LiteralPath $TaskCard -Raw -Encoding UTF8
if ([string]::IsNullOrWhiteSpace($prompt)) {
    Write-Log -Level ERROR -Message "任务卡内容为空：$TaskCard"
    exit 3
}
Write-Log -Level OK -Message "任务卡读取成功（$($prompt.Length) 字符）"

# --- 正式开工：弹出 Agent 状态窗口 ---------------------------------
# 位置很关键：在 $DryRun 提前 return **之后**、claude 循环**之前**。
# 因此只有「真的要去调 claude」时才会弹窗 —— DryRun、.env 缺失、任务卡为空
# 这些提前退出的路径都到不了这里。
Show-AgentStatusWindow

Write-Log -Level STEP -Message "[5/6] 调用 claude -p（最多 $MAX_ATTEMPTS 次，不无限重试）"

$exitCode = 1
for ($attempt = 1; $attempt -le $MAX_ATTEMPTS; $attempt++) {
    Write-Log -Message "—— 第 $attempt / $MAX_ATTEMPTS 次尝试 ——"
    $t0 = Get-Date

    # 通过 stdin 传 prompt，避免 Windows 命令行长度与引号转义问题
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $raw = $prompt | & $claudeExe @claudeArgs 2>&1
        $exitCode = $LASTEXITCODE
    } catch {
        $raw = @("调用 claude 抛出异常：$($_.Exception.Message)")
        $exitCode = 1
    } finally {
        $ErrorActionPreference = $prevEAP
    }

    $text = ($raw | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" }
    }) -join "`n"

    $elapsed = [int]((Get-Date) - $t0).TotalSeconds

    # 原始输出落盘（已脱敏），便于排查
    $outFile = Join-Path $script:LogDir ("claude_output_{0}_attempt{1}.txt" -f $stamp, $attempt)
    try {
        Set-Content -LiteralPath $outFile -Value (Protect-LogText $text) -Encoding UTF8
        Write-Log -Message "本次输出已保存：$outFile"
    } catch {
        Write-Log -Level WARN -Message "输出落盘失败：$($_.Exception.Message)"
    }

    if ($exitCode -eq 0) {
        Write-Log -Level OK -Message "claude 运行成功（耗时 ${elapsed}s）"
        break
    }

    Write-Log -Level ERROR -Message "claude 运行失败，退出码 $exitCode（耗时 ${elapsed}s）"
    if ($text) {
        Write-Log -Level ERROR -Message "错误输出（已脱敏，最多 2000 字符）："
        $snippet = if ($text.Length -gt 2000) { $text.Substring(0, 2000) + ' ...[截断]' } else { $text }
        foreach ($ln in $snippet -split "`n") { Write-Log -Level ERROR -Message "    $ln" }
    }

    if ($attempt -lt $MAX_ATTEMPTS) {
        Write-Log -Level WARN -Message "将在 15 秒后重试一次（这是唯一一次自动重试）"
        Start-Sleep -Seconds 15
    }
}

if ($exitCode -ne 0) {
    Write-Log -Level ERROR -Message "[6/6] 两次尝试均失败，按规则停止（不做无限重试）。"
    Write-Log -Level ERROR -Message "请人工检查上方错误信息与日志：$($script:LogPath)"
    Write-Host ''
    Write-Host "运行失败。日志：$($script:LogPath)" -ForegroundColor Red
    exit 1
}

# --- 7. 产出确认 ---------------------------------------------------
Write-Log -Level STEP -Message "[6/6] 产出确认"

$expected = @(
    # Word 中间底稿（2026-09-16 起不再分「正式版 / 备选版」）
    '本周公司公告汇总.md'
    '本周公司公告来源清单.csv'
    '本周重要资讯.md'
    # PPT 二次筛选结果（每条标 [正式入选] / [备选]）
    'PPT候选内容.md'
    'run_log.md'
    '周报更新说明.md'
    # 以下两份为 Word 交付物，由 脚本\生成Word.py 按 01_模板\公司公告周报模板.docx 渲染。
    # md 仍是 Agent 的中间产物与可追溯底稿，两者都必须产出。
    '本周公司公告周报.docx'
    '本周重要资讯.docx'
)
# 已废止，不再产出：本周公司公告正式版.md / 本周公司公告备选.md / 本周公司公告备选.docx
$missingOut = @()
foreach ($f in $expected) {
    $p = Join-Path $reportDir $f
    if (Test-Path -LiteralPath $p) {
        $size = (Get-Item -LiteralPath $p).Length
        Write-Log -Level OK -Message "  ✔ $f  ($size bytes)"
    } else {
        Write-Log -Level WARN -Message "  ✘ 未生成：$f"
        $missingOut += $f
    }
}
if ($missingOut.Count -gt 0) {
    Write-Log -Level WARN -Message "有 $($missingOut.Count) 个预期文件未生成，请在 05_输出\$Week\run_log.md 中核对原因。"
}

# PPT 页（步骤 8 产出）；依赖状态已在 [3/6] 依赖检查中记录
# 文件名与 脚本\更新PPT.py 保持一致：房地产周报_YYYY.MM.DD-YYYY.MM.DD.pptx
$pptName = '房地产周报_{0}-{1}.pptx' -f $StartDate.ToString('yyyy.MM.dd'), $EndDate.ToString('yyyy.MM.dd')
$pptOut = Join-Path $reportDir $pptName
if (Test-Path -LiteralPath $pptOut) {
    Write-Log -Level OK -Message "  ✔ $pptName  ($((Get-Item -LiteralPath $pptOut).Length) bytes)"
} else {
    Write-Log -Level WARN -Message "  ✘ 未生成：$pptName（原因见 $reportDir\run_log.md）"
}

$total = [int]((Get-Date) - $startTime).TotalSeconds
Write-Log -Level OK -Message "===== 运行结束，总耗时 ${total}s ====="

Write-Host ''
Write-Host "运行完成。" -ForegroundColor Green
Write-Host "输出目录：$reportDir" -ForegroundColor Green
Write-Host "日志：$($script:LogPath)" -ForegroundColor Green
Write-Host ''
exit 0
