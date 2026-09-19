#Requires -Version 5.1
<#
================================================================================
 地产公司公告Agent.ps1 — 地产公司公告 Agent 唯一主入口
================================================================================
 版本：v2.2 ／ 建立日期：2026-09-18
 由旧版四个脚本（run_weekly.ps1、监听待处理目录.ps1、检查并触发周报.ps1、
 查看Agent状态.ps1）整合而成。旧脚本已归档保留，未被删除。

 用法
   .\地产公司公告Agent.ps1 init                     建立目录、检查依赖与 .env
   .\地产公司公告Agent.ps1 run                      运行本周（自动识别最新共同周次）
   .\地产公司公告Agent.ps1 run -Week 2026-W38       运行指定周次
   .\地产公司公告Agent.ps1 run -DryRun              只做校验，不调用模型
   .\地产公司公告Agent.ps1 watch                    实时监听待处理目录（常驻）
   .\地产公司公告Agent.ps1 status                   打开状态窗口
   .\地产公司公告Agent.ps1 install-watcher          注册登录自启动计划任务（需确认）
   .\地产公司公告Agent.ps1 uninstall-watcher        删除计划任务（需二次确认）
   .\地产公司公告Agent.ps1 build-release            生成可分享发行包

 ★★ 安全承诺（硬约束，改动时不得破坏）★★
   · 监听器与所有本地检查**绝不调用任何大模型**、不产生 API 请求。
   · 只有 `run` 正式开工时才调用 claude。
   · 绝不打印、记录、写死任何 API Key / Token；.env 只在进程内存里解析一次。
   · 只写 04_工作区\、05_输出\、日志\、发行版\；待处理公告与历史资料一律只读。
   · 绝不删除待处理公告、不移动公告、不改模板原件、不动注册表与系统代理。
   · 不自动 pip install / npm install；缺依赖只提示人工安装。
================================================================================
#>
[CmdletBinding()]
param(
    # 命令动词
    [Parameter(Position = 0)]
    [ValidateSet('init', 'run', 'watch', 'status', 'install-watcher',
                 'uninstall-watcher', 'build-release', 'help')]
    [string]$Command = 'help',

    # run：指定周次标签，形如 2026-W38
    [ValidatePattern('^\d{4}-W\d{2}$')]
    [string]$Week,

    # run：报告窗口末日（应为周五）
    [datetime]$EndDate,

    # run：只做校验与目录准备，不调用模型
    [switch]$DryRun,

    # watch：稳定性判定所需的「连续无变化」分钟数
    [int]$StableMinutes = 5,

    # watch：稳定扫描间隔（秒）
    [int]$ScanIntervalSeconds = 30,

    # watch：同一周次自动尝试上限
    [int]$MaxRuns = 2,

    # watch：失败后的冷却分钟数
    [int]$FailureCooldownMinutes = 5,

    # watch：阻塞明细的心跳分钟数（避免每 30 秒刷屏）
    [int]$HeartbeatMinutes = 10,

    # watch：跳过启动时的初始扫描
    [switch]$NoInitialScan,

    # watch：跑够 N 轮就退出（测试用；0 = 常驻）
    [int]$MaxCycles = 0,

    # install-watcher / uninstall-watcher：跳过交互式确认（供无人值守调用）
    # 为什么需要它：注册计划任务必须提权，而提权窗口里起 PowerShell 时常常没有
    # 可用的标准输入，Read-Host 会**永久阻塞**——表现就是「卡住不返回、任务也没注册」。
    # 有了 -Yes 就能把「提权」和「确认」分开：由用户在提权窗口里执行本命令。
    [switch]$Yes,

    # status / run：不弹窗（内部使用）
    [switch]$NoWindow,

    # 内部使用：以窗口模式运行（由 Show-AgentStatus 以独立进程方式拉起）
    [switch]$Window
)

$ErrorActionPreference = 'Stop'
# 刻意不开 Set-StrictMode：周次状态是哈希表，StrictMode 2.0 下读一个
# 尚未出现的键会直接抛「属性不存在」，让一个无人值守的监听器半夜崩掉。
# 代价是拼错的变量名不再报错，靠 §三十二 的静态检查补。

# ================================================================== 路径常量
# 全部相对脚本自身位置推导，不写死用户名、不写死绝对目录。
$script:AgentRoot   = $PSScriptRoot
if ([string]::IsNullOrEmpty($script:AgentRoot)) {
    $script:AgentRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$script:ScriptPath  = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($script:ScriptPath)) {
    $script:ScriptPath = Join-Path $script:AgentRoot '地产公司公告Agent.ps1'
}

$script:ConfigDir   = Join-Path $script:AgentRoot '00_配置'
$script:AgentConfig = Join-Path $script:ConfigDir 'Agent配置.json'
$script:PersonalCfg = Join-Path $script:ConfigDir '个人配置.json'
$script:RulesMd     = Join-Path $script:ConfigDir '规则.md'
$script:TaskCard    = Join-Path $script:ConfigDir '任务卡.txt'

$script:InternalDir = Join-Path $script:AgentRoot 'internal'
$script:AgentCorePy = Join-Path $script:InternalDir 'agent_core.py'

$script:TemplateDir = Join-Path $script:AgentRoot '01_模板'
$script:InputRoot   = Join-Path $script:AgentRoot '02_待处理'
$script:OutRoot     = Join-Path $script:AgentRoot '05_输出'
$script:WorkRoot    = Join-Path $script:AgentRoot '04_工作区'
$script:LogDir      = Join-Path $script:AgentRoot '日志'
$script:ReleaseDir  = Join-Path $script:AgentRoot '发行版'
$script:PersonalDir = Join-Path $script:AgentRoot '03_个人资料'

$script:LockDir     = Join-Path $script:WorkRoot 'locks'
$script:EnvPath     = Join-Path $script:AgentRoot '.env'
$script:EnvExample  = Join-Path $script:AgentRoot '.env.example'

$script:WatcherLock = Join-Path $script:LockDir 'folder_watcher.lock'
$script:StatusLock  = Join-Path $script:LockDir 'agent_status_window.lock'

$script:INPUT_PARTS = @('A股公告', '港股公告', '研报', '新闻资讯')
$script:WEEK_DIR_RE = '^W(\d{1,2})公司公告$'

$script:AGENT_VERSION = 'v2.2'
$script:TASK_NAME     = 'RealEstateCompanyAnnouncementAgentWatcher'
$script:OLD_TASK_NAMES = @('RealEstateWeeklyWatcher',          # 旧版周任务
                           'RealEstateWeeklyFolderWatcher')    # 旧版登录常驻监听
$script:STATUS_TITLE  = '地产公司公告 Agent 状态'

$script:MAX_ATTEMPTS  = 2

# 日志文件路径（各入口各自初始化）
$script:LogPath = $null

# ================================================================== 日志
# 内存脱敏表：凡是被登记进来的字面量，写日志前一律替换掉。
$script:Secrets = @()

function Protect-LogText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $out = $Text
    foreach ($s in $script:Secrets) {
        if (-not [string]::IsNullOrEmpty($s) -and $s.Length -ge 6) {
            $out = $out.Replace($s, '****')
        }
    }
    return $out
}

function Protect-Secret {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '(未设置)' }
    if ($Value.Length -le 6) { return '****' }
    return $Value.Substring(0, 6) + '****'
}

function Write-AgentLog {
    param(
        [Parameter(Position = 0)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP', 'SKIP')]
        [string]$Level = 'INFO',
        [string]$Path
    )
    $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, (Protect-LogText $Message)

    switch ($Level) {
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        'SKIP'  { Write-Host $line -ForegroundColor DarkGray }
        default { Write-Host $line }
    }

    $target = if ($Path) { $Path } else { $script:LogPath }
    if ($target) {
        try {
            $dir = Split-Path -Parent $target
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            Add-Content -LiteralPath $target -Value $line -Encoding UTF8 -ErrorAction Stop
        } catch { }
    }
}

function Initialize-AgentLog {
    param([Parameter(Mandatory)][ValidateSet('run', 'watcher', 'agent')][string]$Kind)
    if (-not (Test-Path -LiteralPath $script:LogDir)) {
        New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
    }
    $stamp = (Get-Date).ToString('yyyyMMdd')
    if ($Kind -eq 'run') {
        $script:LogPath = Join-Path $script:LogDir ('run_{0}-{1}.log' -f $stamp, (Get-Date).ToString('HHmmss'))
    } elseif ($Kind -eq 'watcher') {
        $script:LogPath = Join-Path $script:LogDir ("watcher_{0}.log" -f $stamp)
    } else {
        $script:LogPath = Join-Path $script:LogDir ("agent_{0}.log" -f $stamp)
    }
    return $script:LogPath
}

# ================================================================== 周次工具
function Get-IsoWeek {
    param([Parameter(Mandatory)][datetime]$Date)
    $cal = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
    $dow = [int]$Date.DayOfWeek
    if ($dow -eq 0) { $dow = 7 }                      # 周日算作 7
    $thursday = $Date.AddDays(4 - $dow)               # ISO 周归属看周四
    $year = $thursday.Year
    $jan1 = New-Object datetime($year, 1, 1)
    $week = [int][math]::Floor((($thursday - $jan1).Days) / 7) + 1
    return [pscustomobject]@{ IsoYear = $year; IsoWeek = $week }
}

function Format-WeekTag {
    param([Parameter(Mandatory)][int]$Number)
    # 待处理目录只写了周号（W38公司公告），没写年份。归属规则：
    # 以今天的 ISO 年 ISO 周为基准，周号 <= 当前周号 → 属于本年；否则属于上一年。
    # 这样跨年时（1 月还在收 W52）也能算对，不必真去读目录里文件的日期。
    $today = Get-IsoWeek -Date (Get-Date)
    $year = $today.IsoYear
    if ($Number -gt $today.IsoWeek) { $year = $year - 1 }
    return '{0}-W{1:d2}' -f $year, $Number
}

function Get-WeekNumberFromPath {
    # 只认目录名 W数字公司公告；文件名带扩展名，匹配不上，天然被忽略
    param([string]$FullPath)
    if ([string]::IsNullOrWhiteSpace($FullPath)) { return $null }
    $parts = $FullPath -split '[\\/]'
    foreach ($p in $parts) {
        $m = [regex]::Match($p, $script:WEEK_DIR_RE)
        if ($m.Success) { return [int]$m.Groups[1].Value }
    }
    return $null
}

function Get-InputWeeks {
    # 扫描待处理，返回 周号 -> 已存在的分区列表
    $found = @{}
    foreach ($part in $script:INPUT_PARTS) {
        $base = Join-Path $script:InputRoot $part
        if (-not (Test-Path -LiteralPath $base -PathType Container)) { continue }
        foreach ($d in @(Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue)) {
            $m = [regex]::Match($d.Name, $script:WEEK_DIR_RE)
            if (-not $m.Success) { continue }
            $n = [int]$m.Groups[1].Value
            if (-not $found.ContainsKey($n)) { $found[$n] = New-Object System.Collections.ArrayList }
            [void]$found[$n].Add($part)
        }
    }
    return $found
}

function Get-NewestCommonWeekNumber {
    $found = Get-InputWeeks
    if ($found.Count -eq 0) { return $null }
    $both = @($found.Keys | Where-Object {
        $found[$_].Contains('A股公告') -and $found[$_].Contains('港股公告')
    } | Sort-Object)
    if ($both.Count -gt 0) { return [int]$both[-1] }
    return [int](@($found.Keys | Sort-Object)[-1])
}

function Get-WeekPaths {
    param([Parameter(Mandatory)][string]$WeekTag)
    $outDir = Join-Path $script:OutRoot $WeekTag
    return [pscustomobject]@{
        Tag         = $WeekTag
        ReportDir   = $outDir
        DoneFlag    = Join-Path $outDir 'done.flag'
        RunningLock = Join-Path $script:LockDir ("$WeekTag.running")
        ManualLock  = Join-Path $script:LockDir ("$WeekTag.manual")
        AttemptFile = Join-Path $script:LockDir ("$WeekTag.attempts")
    }
}

function Get-WeekRange {
    param([Parameter(Mandatory)][string]$WeekTag)
    $m = [regex]::Match($WeekTag, '^(\d{4})-W(\d{1,2})$')
    if (-not $m.Success) { throw "周次格式应为 YYYY-Wxx：$WeekTag" }
    $y = [int]$m.Groups[1].Value
    $w = [int]$m.Groups[2].Value
    $jan4 = New-Object datetime($y, 1, 4)
    $mon = $jan4.AddDays(-(([int]$jan4.DayOfWeek + 6) % 7)).AddDays(($w - 1) * 7)
    return [pscustomobject]@{ Start = $mon; End = $mon.AddDays(6) }
}

function Resolve-TargetWeek {
    # 计算「本周目标周次」：报告窗口 = 周六 → 周五
    if ($Week) { return $Week }

    if (-not $PSBoundParameters.ContainsKey('EndDate')) {
        $today = (Get-Date).Date
        $dow = [int]$today.DayOfWeek                  # 周五 = 5
        $back = ($dow - 5 + 7) % 7
        $EndDate = if ($back -eq 0) { $today } else { $today.AddDays(-$back) }
    }
    $End = $EndDate.Date
    $iso = Get-IsoWeek -Date $End
    return '{0}-W{1:d2}' -f $iso.IsoYear, $iso.IsoWeek
}

function Test-WeekReady {
    # 输入是否到齐：A股/港股 两侧都存在该周次目录且非空
    param([Parameter(Mandatory)][int]$Number)
    $a = Join-Path (Join-Path $script:InputRoot 'A股公告') ('W{0:d2}公司公告' -f $Number)
    $hk = Join-Path (Join-Path $script:InputRoot '港股公告') ('W{0:d2}公司公告' -f $Number)

    $aOk = (Test-Path -LiteralPath $a -PathType Container) -and (@(Get-ChildItem -LiteralPath $a -File -ErrorAction SilentlyContinue).Count -gt 0)
    $hkOk = (Test-Path -LiteralPath $hk -PathType Container) -and (@(Get-ChildItem -LiteralPath $hk -File -ErrorAction SilentlyContinue).Count -gt 0)

    return [pscustomobject]@{
        Number = $Number
        PathA  = $a
        PathHK = $hk
        HasA   = (Test-Path -LiteralPath $a -PathType Container)
        HasHK  = (Test-Path -LiteralPath $hk -PathType Container)
        NonEmptyA = $aOk
        NonEmptyHK = $hkOk
        Ready  = ($aOk -and $hkOk)
    }
}

# ================================================================== 环境检查
function Test-ProcessAlive {
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return $false }
    try { $null = Get-Process -Id $ProcessId -ErrorAction Stop; return $true } catch { return $false }
}

function Read-DotEnv {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw ".env 不存在：$Path" }
    $vars = @{}
    foreach ($raw in [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)) {
        $line = $raw.Trim().TrimStart([char]0xFEFF)
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

function Resolve-ClaudeExe {
    $cmd = Get-Command claude -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }
    $src = $cmd.Source
    if ([string]::IsNullOrEmpty($src)) { return $null }
    if ($src -notmatch '\.(ps1|cmd|bat)$') { return $src }

    $npmDir = Split-Path -Parent $src
    foreach ($c in @(
        (Join-Path $npmDir 'node_modules\@anthropic-ai\claude-code\bin\claude.exe')
        (Join-Path $npmDir 'claude.exe')
        (Join-Path $npmDir 'claude.cmd')
    )) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return $src
}

function Invoke-AgentPython {
    # 调用 internal\agent_core.py，返回 @{ Code; Lines }
    param([Parameter(Mandatory)][string[]]$Arguments)
    $py = Get-Command python -ErrorAction SilentlyContinue
    if (-not $py) { return [pscustomobject]@{ Code = 127; Lines = @('未找到 python') } }
    if (-not (Test-Path -LiteralPath $script:AgentCorePy)) {
        return [pscustomobject]@{ Code = 127; Lines = @("未找到 $script:AgentCorePy") }
    }

    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & python $script:AgentCorePy @Arguments 2>&1
        $code = $LASTEXITCODE
    } catch {
        $out = @("调用失败：$($_.Exception.Message)")
        $code = 1
    } finally {
        $ErrorActionPreference = $prev
    }
    $lines = @($out | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" }
    })
    return [pscustomobject]@{ Code = $code; Lines = $lines }
}

function Test-AgentEnvironment {
    <#
      返回 @{ Ok; Problems; Missing }。
      只报告，不修复、不安装、不改动任何文件。
    #>
    $problems = New-Object System.Collections.ArrayList
    $missing = New-Object System.Collections.ArrayList

    # 1. 必需目录
    foreach ($p in @($script:InputRoot, $script:WorkRoot, $script:OutRoot,
                     $script:LogDir, $script:TemplateDir, $script:ConfigDir)) {
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            [void]$problems.Add("缺失目录：$p")
        }
    }

    # 2. 模板
    foreach ($t in @('公司公告周报模板.docx', '房地产周报PPT模板.pptx')) {
        $p = Join-Path $script:TemplateDir $t
        if (-not (Test-Path -LiteralPath $p)) { [void]$problems.Add("缺失模板：$t") }
    }

    # 3. Python 依赖（只检查，不安装）
    $r = Invoke-AgentPython -Arguments @('deps')
    foreach ($ln in $r.Lines) {
        if ($ln -match '^\s*MISS\s+(\S+)') { [void]$missing.Add($Matches[1]) }
    }
    if ($r.Code -eq 127) { [void]$problems.Add("无法调用 Python 或缺少 $script:AgentCorePy") }

    # 4. .env
    if (-not (Test-Path -LiteralPath $script:EnvPath)) {
        [void]$problems.Add("未发现 .env —— 请复制 .env.example 为 .env，并填写自己的 API Key。")
    }

    # 5. 任务卡
    if (-not (Test-Path -LiteralPath $script:TaskCard)) {
        [void]$problems.Add("缺失任务卡：$script:TaskCard")
    }

    return [pscustomobject]@{
        Ok       = ($problems.Count -eq 0 -and $missing.Count -eq 0)
        Problems = @($problems)
        Missing  = @($missing)
    }
}

function Initialize-Agent {
    <#
      建立目录、检查依赖与 .env。只创建目录，绝不删除或覆盖任何东西。
    #>
    Write-AgentLog -Level STEP -Message '===== 初始化 地产公司公告 Agent ====='
    Write-AgentLog -Level STEP -Message '[1/3] 建立目录'

    foreach ($p in @($script:ConfigDir, $script:InputRoot, $script:OutRoot,
                     $script:WorkRoot, $script:LogDir, $script:ReleaseDir,
                     $script:PersonalDir, $script:TemplateDir, $script:LockDir,
                     (Join-Path $script:WorkRoot '盘点'))) {
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
            Write-AgentLog -Message "  已创建：$p"
        }
    }
    foreach ($part in $script:INPUT_PARTS) {
        $p = Join-Path $script:InputRoot $part
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
            Write-AgentLog -Message "  已创建：$p"
        }
    }
    Write-AgentLog -Level OK -Message '  目录就绪'

    Write-AgentLog -Level STEP -Message '[2/3] 依赖检查（只检查，不安装）'
    $r = Invoke-AgentPython -Arguments @('deps')
    foreach ($ln in $r.Lines) { Write-AgentLog -Message "  $ln" }
    if ($r.Code -ne 0) {
        Write-AgentLog -Level WARN -Message '  有依赖缺失。请自行执行：pip install -r requirements.txt'
        Write-AgentLog -Level WARN -Message '  本 Agent 不会自动安装任何软件包。'
    } else {
        Write-AgentLog -Level OK -Message '  依赖齐全'
    }

    Write-AgentLog -Level STEP -Message '[3/3] API 配置'
    if (-not (Test-Path -LiteralPath $script:EnvPath)) {
        Write-AgentLog -Level WARN -Message '  未发现 .env'
        Write-AgentLog -Level WARN -Message '  请复制 .env.example 为 .env，并填写自己的 API Key。'
    } else {
        try {
            $vars = Read-DotEnv -Path $script:EnvPath
            Write-AgentLog -Level OK -Message ("  已发现 .env，共 {0} 个配置项（内容不打印）" -f $vars.Count)
            $empty = @($vars.Keys | Where-Object { [string]::IsNullOrWhiteSpace($vars[$_]) })
            if ($empty.Count -gt 0) {
                Write-AgentLog -Level WARN -Message ("  以下项尚未填写：{0}" -f ($empty -join '、'))
            }
        } catch {
            Write-AgentLog -Level WARN -Message "  .env 解析失败：$($_.Exception.Message)"
        }
    }

    Write-AgentLog -Level OK -Message '初始化完成。'
    return 0
}

# ================================================================== run
function Test-InputFiles {
    # 只读扫描本周两侧输入；绝不删除、不移动、不改名
    param([Parameter(Mandatory)][string]$WeekTag)

    $range = Get-WeekRange -WeekTag $WeekTag
    $m = [regex]::Match($WeekTag, '^(\d{4})-W(\d{1,2})$')
    $num = [int]$m.Groups[2].Value

    $dirA = Join-Path (Join-Path $script:InputRoot 'A股公告') ('W{0:d2}公司公告' -f $num)
    $dirHK = Join-Path (Join-Path $script:InputRoot '港股公告') ('W{0:d2}公司公告' -f $num)

    $aFiles = @()
    $hkFiles = @()
    if (Test-Path -LiteralPath $dirA -PathType Container) {
        $aFiles = @(Get-ChildItem -LiteralPath $dirA -File -ErrorAction SilentlyContinue |
                    Where-Object { -not $_.Name.StartsWith('~$') })
    }
    if (Test-Path -LiteralPath $dirHK -PathType Container) {
        $hkFiles = @(Get-ChildItem -LiteralPath $dirHK -File -ErrorAction SilentlyContinue |
                     Where-Object { -not $_.Name.StartsWith('~$') })
    }

    return [pscustomobject]@{
        WeekTag = $WeekTag
        Number  = $num
        DirA    = $dirA
        DirHK   = $dirHK
        CountA  = $aFiles.Count
        CountHK = $hkFiles.Count
        Start   = $range.Start
        End     = $range.End
        Ready   = (Test-WeekReady -Number $num)
    }
}

function Get-CrossMarketDuplicates {
    <#
      跨市场重复检测。同一份公告在 A股、港股 两侧各投一次是常见情形。

      规则（照搬旧版口径）：以 **A股目录为基准**，港股目录中
      「同名 → 同大小 → 同 SHA256」的文件判定为跨市场重复，
      **不计入港股有效文件数**；A股一侧原样保留，数量不变。

      ★ 只在处理阶段标记忽略；磁盘文件一个都不动：不删、不移、不改名。
      返回 @{ Dups; NameOnly; AEffective; HKEffective }。
    #>
    param([Parameter(Mandatory)]$Info)

    $dups = New-Object System.Collections.ArrayList
    $nameOnly = New-Object System.Collections.ArrayList

    $aFiles = @()
    $hkFiles = @()
    if (Test-Path -LiteralPath $Info.DirA -PathType Container) {
        $aFiles = @(Get-ChildItem -LiteralPath $Info.DirA -File -ErrorAction SilentlyContinue |
                    Where-Object { -not $_.Name.StartsWith('~$') })
    }
    if (Test-Path -LiteralPath $Info.DirHK -PathType Container) {
        $hkFiles = @(Get-ChildItem -LiteralPath $Info.DirHK -File -ErrorAction SilentlyContinue |
                     Where-Object { -not $_.Name.StartsWith('~$') })
    }

    if ($aFiles.Count -gt 0 -and $hkFiles.Count -gt 0) {
        $aIndex = @{}
        foreach ($f in $aFiles) {
            if (-not $aIndex.ContainsKey($f.Name)) { $aIndex[$f.Name] = $f }
        }
        foreach ($h in $hkFiles) {
            if (-not $aIndex.ContainsKey($h.Name)) { continue }        # 1) 文件名
            $a = $aIndex[$h.Name]
            if ($a.Length -ne $h.Length) {                              # 2) 文件大小
                [void]$nameOnly.Add("$($h.Name)（同名但大小不同：A股 $($a.Length) B / 港股 $($h.Length) B）")
                continue
            }
            $shaA = $null; $shaH = $null                                # 3) 哈希（可行时）
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
            [void]$dups.Add([pscustomobject]@{
                '忽略文件路径' = $h.FullName
                '对应A股文件'  = $a.FullName
                '文件名'       = $h.Name
                '字节数'       = $h.Length
                'sha256'       = $(if ($shaH) { $shaH } else { '' })
                '判定依据'     = $basis
                '处理方式'     = '忽略：跨市场重复文件（不删除、不移动、不改名）'
            })
        }
    }

    $hkEffective = $hkFiles.Count - $dups.Count
    if ($hkEffective -lt 0) { $hkEffective = 0 }

    return [pscustomobject]@{
        Dups        = $dups
        NameOnly    = $nameOnly
        AEffective  = $aFiles.Count
        HKEffective = $hkEffective
    }
}


function Start-WeeklyRun {
    <#
      正式运行：加载 .env → 计算周次 → 前置校验 → 调 claude -p（最多 2 次）→ 产出确认。
      只有本函数会调用模型。
    #>
    param([switch]$IsDryRun)

    $startTime = Get-Date
    $null = Initialize-AgentLog -Kind 'run'

    Write-Host ''
    Write-Host '==================== 地产公司公告 Agent · 每周运行 ====================' -ForegroundColor Cyan
    Write-Host ("工作区：{0}" -f $script:AgentRoot)

    Set-Location -LiteralPath $script:AgentRoot
    Write-AgentLog -Level STEP -Message '===== 地产公司公告 Agent 运行开始 ====='
    Write-AgentLog -Message "工作区：$script:AgentRoot"
    Write-AgentLog -Message "日志文件：$script:LogPath"

    if ($IsDryRun) {
        Write-AgentLog -Level WARN -Message 'DryRun 模式：只做校验与目录准备，不调用任何模型、不产生 API 请求。'
    }

    # --- 1. .env ----------------------------------------------------
    Write-AgentLog -Level STEP -Message '[1/6] 读取 .env'
    if (-not (Test-Path -LiteralPath $script:EnvPath)) {
        Write-AgentLog -Level ERROR -Message "未发现 .env：$script:EnvPath"
        Write-AgentLog -Level WARN  -Message '请复制 .env.example 为 .env，并填写自己的 API Key。'
        return 2
    }
    try {
        $envVars = Read-DotEnv -Path $script:EnvPath
    } catch {
        Write-AgentLog -Level ERROR -Message "读取 .env 失败：$($_.Exception.Message)"
        return 2
    }

    $required = @('ANTHROPIC_BASE_URL', 'ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_MODEL')
    $missEnv = @($required | Where-Object {
        -not $envVars.ContainsKey($_) -or [string]::IsNullOrWhiteSpace($envVars[$_]) })
    if ($missEnv.Count -gt 0) {
        Write-AgentLog -Level ERROR -Message ("缺少必需变量：{0}" -f ($missEnv -join ', '))
        return 2
    }

    # 密钥只进内存脱敏表与子进程环境，绝不落盘、绝不打印
    $script:Secrets = @($envVars['ANTHROPIC_AUTH_TOKEN'])
    if (-not $IsDryRun) {
        $env:ANTHROPIC_BASE_URL   = $envVars['ANTHROPIC_BASE_URL']
        $env:ANTHROPIC_AUTH_TOKEN = $envVars['ANTHROPIC_AUTH_TOKEN']
        $env:ANTHROPIC_MODEL      = $envVars['ANTHROPIC_MODEL']
    }
    Write-AgentLog -Level OK -Message "ANTHROPIC_BASE_URL = $($envVars['ANTHROPIC_BASE_URL'])"
    Write-AgentLog -Level OK -Message "ANTHROPIC_AUTH_TOKEN = $(Protect-Secret $envVars['ANTHROPIC_AUTH_TOKEN'])  （已遮蔽，永不完整输出）"
    Write-AgentLog -Level OK -Message "ANTHROPIC_MODEL = $($envVars['ANTHROPIC_MODEL'])"

    # --- 2. 周次与输出目录 -------------------------------------------
    Write-AgentLog -Level STEP -Message '[2/6] 计算周次与输出目录'
    $weekTag = Resolve-TargetWeek
    $info = Test-InputFiles -WeekTag $weekTag
    $reportDir = Join-Path $script:OutRoot $weekTag

    Write-AgentLog -Message ("报告窗口：{0}（周六） → {1}（周五）" -f `
        $info.Start.ToString('yyyy-MM-dd'), $info.End.ToString('yyyy-MM-dd'))
    Write-AgentLog -Message "周次标签：$weekTag"

    if (-not (Test-Path -LiteralPath $reportDir)) {
        New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
        Write-AgentLog -Level OK -Message "已建立输出目录：$reportDir"
    } else {
        Write-AgentLog -Message "输出目录已存在：$reportDir"
    }

    # done.flag 语义（规范 §十三）：
    #   · 对**自动监听**而言，done.flag 是「本周已完成，不要重复触发」的硬性拦阻；
    #   · 对**手动 run** 而言，它只是提示，不是拦阻。用户明确指定或自动识别到
    #     同一周次时，允许重新生成并覆盖该周次结果。
    # 这里只做提示，不 return —— 覆盖与否由用户「是否执行 run」决定。
    $doneFlag = Join-Path $reportDir 'done.flag'
    if (Test-Path -LiteralPath $doneFlag) {
        Write-AgentLog -Level WARN -Message "本周已完成过（done.flag 存在）：$doneFlag"
        if ($IsDryRun) {
            Write-AgentLog -Level WARN -Message '  这是 DryRun：只做校验，不会写入或覆盖任何产出。'
        } else {
            Write-AgentLog -Level WARN -Message '  本次会重新生成并覆盖该周次已有结果；待处理目录中的原始公告不会被改动。'
            Write-Host ("提示：{0} 已完成过，本次运行会覆盖该周次已有结果。" -f $weekTag) -ForegroundColor Yellow
        }
    }

    $env:REPORT_WEEK  = $weekTag
    $env:REPORT_DIR   = $reportDir
    $env:REPORT_START = $info.Start.ToString('yyyy-MM-dd')
    $env:REPORT_END   = $info.End.ToString('yyyy-MM-dd')

    Write-AgentLog -Message ("本周输入：A股 {0} 个（{1}）" -f $info.CountA, $info.DirA)
    Write-AgentLog -Message ("          港股 {0} 个（{1}）" -f $info.CountHK, $info.DirHK)

    # Test-WeekReady：A股/港股 两侧该周次目录都存在且非空
    $rd = $info.Ready
    Write-AgentLog -Message ("两边齐备情况：A股 目录存在={0} 非空={1}；港股 目录存在={2} 非空={3}" -f `
        $rd.HasA, $rd.NonEmptyA, $rd.HasHK, $rd.NonEmptyHK)
    if ($rd.Ready) {
        Write-AgentLog -Level OK -Message '  → 两侧输入齐备'
    } else {
        Write-AgentLog -Level WARN -Message '  → 两侧输入不齐备，本次产出会不完整（不伪造内容）。'
    }

    # --- 跨市场重复检测（只标记忽略，绝不删除 / 移动 / 改名）---------
    $dup = Get-CrossMarketDuplicates -Info $info
    $dupCsv = Join-Path $reportDir '跨市场重复清单.csv'
    if ($dup.Dups.Count -gt 0) {
        try {
            $dup.Dups | Export-Csv -LiteralPath $dupCsv -NoTypeInformation -Encoding UTF8
            Write-AgentLog -Level WARN -Message ("跨市场重复：港股目录中 {0} 个文件与 A股目录完全重复 → 【忽略：跨市场重复文件】" -f $dup.Dups.Count)
            Write-AgentLog -Level WARN -Message ("  仅在处理阶段忽略，磁盘文件不删除、不移动；清单：{0}" -f $dupCsv)
        } catch {
            Write-AgentLog -Level WARN -Message "跨市场重复清单落盘失败：$($_.Exception.Message)"
        }
    } elseif ($info.CountA -gt 0 -and $info.CountHK -gt 0) {
        Write-AgentLog -Level OK -Message '跨市场重复：未发现重复文件'
    }
    if ($dup.NameOnly.Count -gt 0) {
        Write-AgentLog -Level WARN -Message ("同名但未判定为完全重复（照常处理）：{0} 个" -f $dup.NameOnly.Count)
        foreach ($x in $dup.NameOnly) { Write-AgentLog -Message "    · $x" }
    }
    Write-AgentLog -Level OK -Message ("本周有效输入：A股 {0} 个 / 港股 {1} 个 / 忽略重复 {2} 个" -f `
        $dup.AEffective, $dup.HKEffective, $dup.Dups.Count)

    # 显式声明被忽略的其他周次目录，证明未把上周/下周公告混入本周
    foreach ($mkt in @('A股公告', '港股公告')) {
        $parent = Join-Path $script:InputRoot $mkt
        if (-not (Test-Path -LiteralPath $parent -PathType Container)) { continue }
        $want = 'W{0:d2}公司公告' -f $info.Number
        $others = @(Get-ChildItem -LiteralPath $parent -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -ne $want })
        if ($others.Count -gt 0) {
            Write-AgentLog -Message ("  · 已忽略其他周次目录（{0}）：{1}" -f `
                $mkt, (($others | ForEach-Object { $_.Name }) -join '、'))
        }
    }

    # 交给子进程 / claude 的输入位置一律走环境变量，不写死在提示词里
    $env:REPORT_IN_A = $info.DirA
    $env:REPORT_IN_HK = $info.DirHK
    foreach ($pair in @(@{ K = 'REPORT_IN_NEWS'; P = '新闻资讯' },
                        @{ K = 'REPORT_IN_REPORT'; P = '研报' })) {
        [System.Environment]::SetEnvironmentVariable($pair.K, (Join-Path $script:InputRoot $pair.P), 'Process')
    }
    [System.Environment]::SetEnvironmentVariable('REPORT_A_EFFECTIVE', "$($dup.AEffective)", 'Process')
    [System.Environment]::SetEnvironmentVariable('REPORT_HK_EFFECTIVE', "$($dup.HKEffective)", 'Process')
    [System.Environment]::SetEnvironmentVariable('REPORT_DUP_COUNT', "$($dup.Dups.Count)", 'Process')
    [System.Environment]::SetEnvironmentVariable('REPORT_DUP_CSV',
        $(if ($dup.Dups.Count -gt 0) { $dupCsv } else { '' }), 'Process')

    # --- 3. 前置校验 -------------------------------------------------
    Write-AgentLog -Level STEP -Message '[3/6] 前置校验'

    $envchk = Test-AgentEnvironment
    if ($envchk.Ok) {
        Write-AgentLog -Level OK -Message '  运行环境自检通过'
    } else {
        foreach ($p in $envchk.Problems) { Write-AgentLog -Level WARN -Message "  ! $p" }
    }
    if ($envchk.Missing.Count -gt 0) {
        Write-AgentLog -Level WARN -Message ("  缺少 Python 依赖：{0}" -f ($envchk.Missing -join '、'))
        Write-AgentLog -Level WARN -Message '  请自行执行 pip install -r requirements.txt（本 Agent 不自动安装）。'
    }
    foreach ($rel in @('00_配置\规则.md', '00_配置\Agent配置.json',
                       '01_模板\公司公告周报模板.docx', '01_模板\房地产周报PPT模板.pptx')) {
        $full = Join-Path $script:AgentRoot $rel
        if (Test-Path -LiteralPath $full) {
            Write-AgentLog -Level OK -Message "  ✔ $rel"
        } else {
            Write-AgentLog -Level WARN -Message "  ✘ 缺失：$rel （将记入日志，不伪造内容）"
        }
    }

    if (-not (Test-Path -LiteralPath $script:TaskCard)) {
        Write-AgentLog -Level ERROR -Message "任务卡不存在：$script:TaskCard"
        return 3
    }
    Write-AgentLog -Level OK -Message "任务卡：$script:TaskCard"

    $gitignore = Join-Path $script:AgentRoot '.gitignore'
    if (Test-Path -LiteralPath $gitignore) {
        $gi = Get-Content -LiteralPath $gitignore -Raw -Encoding UTF8
        if ($gi -match '(?m)^\s*\.env\s*$') {
            Write-AgentLog -Level OK -Message '.gitignore 已忽略 .env'
        } else {
            Write-AgentLog -Level WARN -Message '.gitignore 未发现独立的 .env 规则，请检查'
        }
    }

    $claudeExe = Resolve-ClaudeExe
    if (-not $claudeExe) {
        Write-AgentLog -Level ERROR -Message '未找到 claude 命令，请确认 Claude Code CLI 已安装并在 PATH 中'
        return 4
    }
    Write-AgentLog -Level OK -Message "claude 可执行文件：$claudeExe"

    $depRes = Invoke-AgentPython -Arguments @('deps')
    foreach ($ln in $depRes.Lines) { Write-AgentLog -Message "  依赖 $ln" }
    if ($depRes.Code -ne 0) {
        Write-AgentLog -Level WARN -Message '有依赖缺失，相关产出会被跳过（不自动安装）。'
    }

    if ($IsDryRun) {
        Write-AgentLog -Level OK -Message 'DryRun 到此结束：未调用模型、未产生 API 请求。'
        Write-Host ''
        Write-Host 'DryRun 完成。' -ForegroundColor Green
        Write-Host "输出目录：$reportDir" -ForegroundColor Green
        Write-Host "日志：$script:LogPath" -ForegroundColor Green
        return 0
    }

    # --- 4. 组装调用参数 ---------------------------------------------
    Write-AgentLog -Level STEP -Message '[4/6] 组装 claude 调用参数'
    $prompt = Get-Content -LiteralPath $script:TaskCard -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($prompt)) {
        Write-AgentLog -Level ERROR -Message "任务卡为空：$script:TaskCard"
        return 3
    }
    $claudeArgs = @('-p', '--permission-mode', 'acceptEdits')
    if ($envVars.ContainsKey('ANTHROPIC_MODEL') -and $envVars['ANTHROPIC_MODEL']) {
        $claudeArgs += @('--model', $envVars['ANTHROPIC_MODEL'])
    }
    Write-AgentLog -Level OK -Message "任务卡读取成功（$($prompt.Length) 字符）"

    # --- 正式开工：弹状态窗口 ----------------------------------------
    if (-not $NoWindow) { [void](Show-AgentStatus -Detached) }

    # --- 5. 调用模型（最多 2 次）-------------------------------------
    Write-AgentLog -Level STEP -Message ("[5/6] 调用 claude -p（最多 {0} 次，不无限重试）" -f $script:MAX_ATTEMPTS)

    $exitCode = 1
    $stamp = $startTime.ToString('yyyyMMdd-HHmmss')
    for ($attempt = 1; $attempt -le $script:MAX_ATTEMPTS; $attempt++) {
        Write-AgentLog -Message "—— 第 $attempt / $script:MAX_ATTEMPTS 次尝试 ——"
        $t0 = Get-Date

        $prev = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $raw = $prompt | & $claudeExe @claudeArgs 2>&1
            $exitCode = $LASTEXITCODE
        } catch {
            $raw = @("调用 claude 抛出异常：$($_.Exception.Message)")
            $exitCode = 1
        } finally {
            $ErrorActionPreference = $prev
        }

        $text = ($raw | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" }
        }) -join "`n"

        $elapsed = [int]((Get-Date) - $t0).TotalSeconds

        $outFile = Join-Path $script:LogDir ("claude_output_{0}_attempt{1}.txt" -f $stamp, $attempt)
        try {
            Set-Content -LiteralPath $outFile -Value (Protect-LogText $text) -Encoding UTF8
            Write-AgentLog -Message "本次输出已保存：$outFile"
        } catch {
            Write-AgentLog -Level WARN -Message "输出落盘失败：$($_.Exception.Message)"
        }

        if ($exitCode -eq 0) {
            Write-AgentLog -Level OK -Message "claude 运行成功（耗时 ${elapsed}s）"
            break
        }

        Write-AgentLog -Level ERROR -Message "claude 运行失败，退出码 $exitCode（耗时 ${elapsed}s）"
        if ($text) {
            $snippet = if ($text.Length -gt 2000) { $text.Substring(0, 2000) + ' ...[截断]' } else { $text }
            foreach ($ln in ($snippet -split "`n")) {
                Write-AgentLog -Level ERROR -Message "    $ln"
            }
        }
        if ($attempt -lt $script:MAX_ATTEMPTS) {
            Write-AgentLog -Level WARN -Message '将在 15 秒后重试一次（这是唯一一次自动重试）'
            Start-Sleep -Seconds 15
        }
    }

    if ($exitCode -ne 0) {
        Write-AgentLog -Level ERROR -Message '[6/6] 两次尝试均失败，按规则停止（不做无限重试）。'
        Write-AgentLog -Level ERROR -Message "请人工检查上方错误信息与日志：$script:LogPath"
        return 1
    }

    # --- 6. 产出确认 -------------------------------------------------
    Write-AgentLog -Level STEP -Message '[6/6] 产出确认'
    $res = Invoke-AgentPython -Arguments @('validate', '--week', $weekTag)
    foreach ($ln in $res.Lines) { Write-AgentLog -Message "  $ln" }

    if ($res.Code -eq 0) {
        # 全部产出齐全 → 写 done.flag
        $paths = Get-WeekPaths -WeekTag $weekTag
        try {
            Set-Content -LiteralPath $paths.DoneFlag `
                -Value ("DONE $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))") -Encoding ASCII
            Write-AgentLog -Level OK -Message "已写入 done.flag：$($paths.DoneFlag)"
        } catch {
            Write-AgentLog -Level WARN -Message "done.flag 写入失败：$($_.Exception.Message)"
        }
    } else {
        Write-AgentLog -Level WARN -Message '产出不齐全 → 不写 done.flag（原因见上方明细）。'
    }

    $total = [int]((Get-Date) - $startTime).TotalSeconds
    Write-AgentLog -Level OK -Message "===== 运行结束，总耗时 ${total}s ====="
    Write-Host ''
    Write-Host '运行完成。' -ForegroundColor Green
    Write-Host "输出目录：$reportDir" -ForegroundColor Green
    Write-Host "日志：$script:LogPath" -ForegroundColor Green
    return 0
}

# ================================================================== watch
$script:Weeks = @{}
$script:Cycles = 0

function Get-WeekDirInfo {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return [pscustomobject]@{ Exists = $false; Count = 0; Newest = $null }
    }
    $count = 0
    $newest = $null
    foreach ($f in @(Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue)) {
        if ($f.Name.StartsWith('~$')) { continue }
        $count++
        if ($null -eq $newest -or $f.LastWriteTime -gt $newest) { $newest = $f.LastWriteTime }
    }
    foreach ($d in @(Get-ChildItem -LiteralPath $Path -Directory -Recurse -ErrorAction SilentlyContinue)) {
        if ($null -eq $newest -or $d.LastWriteTime -gt $newest) { $newest = $d.LastWriteTime }
    }
    return [pscustomobject]@{ Exists = $true; Count = $count; Newest = $newest }
}

function Format-Age {
    param($Newest)
    if ($null -eq $Newest) { return '（未知）' }
    $span = (Get-Date) - $Newest
    if ($span.TotalMinutes -lt 1) { return ('{0:N0} 秒前' -f $span.TotalSeconds) }
    if ($span.TotalHours -lt 1) { return ('{0:N1} 分钟前' -f $span.TotalMinutes) }
    return ('{0:N1} 小时前' -f $span.TotalHours)
}

function Get-OrCreateWeek {
    param([Parameter(Mandatory)][int]$Number)
    $tag = Format-WeekTag -Number $Number
    if ($script:Weeks.ContainsKey($tag)) { return $script:Weeks[$tag] }

    $entry = @{
        Tag             = $tag
        Number          = $Number
        CreatedAt       = Get-Date
        ALastEventAt    = $null
        HKLastEventAt   = $null
        AEvents         = 0
        HKEvents        = 0
        ACounts         = @{ Created = 0; Changed = 0; Renamed = 0; Deleted = 0 }
        HKCounts        = @{ Created = 0; Changed = 0; Renamed = 0; Deleted = 0 }
        BurstLogged     = $false
        NewestSeen      = $null
        StableLogged    = $false
        LastBlockKey    = ''
        NextHeartbeatAt = $null
        Dispatched      = $false
        ChildProcess    = $null
        ChildStartedAt  = $null
        DeferUntil      = $null
        DeferLogged     = $false
        Terminal        = $false
        TerminalReason  = ''
    }
    $script:Weeks[$tag] = $entry
    return $entry
}

function Set-WeekTerminal {
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][string]$Reason)
    if ($Entry.Terminal) { return }
    $Entry.Terminal = $true
    $Entry.TerminalReason = $Reason
}

function Get-WatcherRunningInfo {
    param([Parameter(Mandatory)]$Paths)
    if (-not (Test-Path -LiteralPath $Paths.RunningLock)) { return $null }
    $procId = 0; $start = ''
    try {
        foreach ($ln in [System.IO.File]::ReadAllLines($Paths.RunningLock)) {
            if ($ln -match '^\s*PID\s*=\s*(\d+)')   { $procId = [int]$Matches[1] }
            if ($ln -match '^\s*START\s*=\s*(.+)$') { $start = $Matches[1].Trim() }
        }
    } catch { }
    return [pscustomobject]@{ ProcessId = $procId; Start = $start; Alive = (Test-ProcessAlive -ProcessId $procId) }
}

function Get-WatcherAttemptCount {
    param([Parameter(Mandatory)]$Paths)
    if (-not (Test-Path -LiteralPath $Paths.AttemptFile)) { return 0 }
    try {
        $n = 0
        if ([int]::TryParse((Get-Content -LiteralPath $Paths.AttemptFile -Raw -ErrorAction Stop).Trim(), [ref]$n)) { return $n }
    } catch { }
    return 0
}

function Enter-WatcherLock {
    if (Test-Path -LiteralPath $script:WatcherLock) {
        $oldPid = 0
        try {
            $m = [regex]::Match((Get-Content -LiteralPath $script:WatcherLock -Raw -ErrorAction Stop), 'PID\s*=\s*(\d+)')
            if ($m.Success) { $oldPid = [int]$m.Groups[1].Value }
        } catch { }

        if ($oldPid -gt 0 -and (Test-ProcessAlive -ProcessId $oldPid)) {
            Write-AgentLog -Level WARN -Message ("检测到监听器已在运行（PID {0}）→ 本实例退出，不重复监听。" -f $oldPid)
            Write-Host ("监听器已在运行（PID {0}），本实例直接退出。" -f $oldPid) -ForegroundColor Yellow
            Write-Host ("锁文件：{0}" -f $script:WatcherLock) -ForegroundColor DarkGray
            return $false
        }
        Write-AgentLog -Level WARN -Message ("发现陈旧监听锁（PID {0} 已不存在）→ 清理后重新启动。" -f $oldPid)
        try { Remove-Item -LiteralPath $script:WatcherLock -Force -ErrorAction Stop } catch {
            Write-AgentLog -Level ERROR -Message ("陈旧锁清理失败：{0}" -f $_.Exception.Message)
            return $false
        }
    }

    try {
        if (-not (Test-Path -LiteralPath $script:LockDir)) {
            New-Item -ItemType Directory -Path $script:LockDir -Force | Out-Null
        }
        $body = @(
            "PID=$PID"
            "START=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
            "SCRIPT=地产公司公告Agent.ps1 watch"
            "VERSION=$script:AGENT_VERSION"
        ) -join "`r`n"
        Set-Content -LiteralPath $script:WatcherLock -Value $body -Encoding UTF8
        Write-AgentLog -Message ("已写入监听器锁：{0}（PID={1}）" -f $script:WatcherLock, $PID)
        return $true
    } catch {
        Write-AgentLog -Level ERROR -Message ("写入监听器锁失败：{0}" -f $_.Exception.Message)
        return $false
    }
}

function Exit-WatcherLock {
    if (-not (Test-Path -LiteralPath $script:WatcherLock)) { return }
    $owner = 0
    try {
        $m = [regex]::Match((Get-Content -LiteralPath $script:WatcherLock -Raw -ErrorAction Stop), 'PID\s*=\s*(\d+)')
        if ($m.Success) { $owner = [int]$m.Groups[1].Value }
    } catch { }
    if ($owner -eq $PID) {
        try { Remove-Item -LiteralPath $script:WatcherLock -Force -ErrorAction Stop } catch { }
    }
}

function Register-AgentChange {
    param(
        [Parameter(Mandatory)][string]$Market,
        [Parameter(Mandatory)][int]$Number,
        [Parameter(Mandatory)][string]$ChangeType,
        [string]$Path,
        [string]$OldPath
    )
    $entry = Get-OrCreateWeek -Number $Number

    if ($entry.Terminal) {
        if (-not $entry.ContainsKey('TerminalNoteLogged')) { $entry['TerminalNoteLogged'] = $false }
        if (-not $entry.TerminalNoteLogged) {
            Write-AgentLog -Level SKIP -Message ("{0} 已处于终态（{1}）→ 本次变化忽略：{2}" -f `
                $entry.Tag, $entry.TerminalReason, $ChangeType)
            $entry.TerminalNoteLogged = $true
        }
        return
    }

    $now = Get-Date
    if ($Market -eq 'A') { $entry.ALastEventAt = $now; $entry.AEvents++ }
    else                 { $entry.HKLastEventAt = $now; $entry.HKEvents++ }
    if (-not $entry["${Market}Counts"].ContainsKey($ChangeType)) { $entry["${Market}Counts"][$ChangeType] = 0 }
    $entry["${Market}Counts"][$ChangeType]++

    # 变化即视为「又重新开始动」，撤销稳定标记，重新计时
    $entry.StableLogged = $false
    $entry.BurstLogged = $true

    # 输入又变了 → 作废上一次的重试冷却（仍受「连续 N 分钟无变化」约束）
    $entry.DeferUntil = $null
    $entry.DeferLogged = $false

    if (($entry.AEvents + $entry.HKEvents) -eq 1) {
        $side = if ($Market -eq 'A') { 'A股' } else { '港股' }
        Write-AgentLog -Message ("检测到 {0} 变化：{1} [{2}] {3}" -f $entry.Tag, $side, $ChangeType, $Path)
        if ($OldPath) { Write-AgentLog -Message ("  原路径：{0}" -f $OldPath) }
        Write-AgentLog -Message ("  → 已更新该周次最后变化时间，进入去抖动等待（{0} 分钟无变化才算稳定）" -f $StableMinutes)
    }
}

function Handle-AgentFileEvent {
    param([Parameter(Mandatory)]$Event)
    $parts = $Event.SourceIdentifier -split '\|'
    $market = $parts[0]
    $kind = $parts[1]
    $ea = $Event.SourceEventArgs

    if ($kind -eq 'Error') {
        Write-AgentLog -Level WARN -Message ("{0} 侧监听缓冲区溢出，可能丢失部分事件；已回退到按磁盘写入时间判定。" -f $market)
        return
    }

    $path = $null
    $oldPath = $null
    if ($ea -is [System.IO.RenamedEventArgs]) {
        $path = $ea.FullPath
        $oldPath = $ea.OldFullPath
    } elseif ($ea -is [System.IO.FileSystemEventArgs]) {
        $path = $ea.FullPath
    }
    if ([string]::IsNullOrWhiteSpace($path)) { return }

    # 改名事件：旧路径属于哪个周次也要一起刷新
    foreach ($p in @($path, $oldPath)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $n = Get-WeekNumberFromPath -FullPath $p
        if ($null -eq $n) { continue }
        Register-AgentChange -Market $market -Number $n -ChangeType $kind -Path $path -OldPath $oldPath
        break
    }
}

function Initialize-PendingWeeks {
    Write-AgentLog -Level STEP -Message '启动初始扫描：检查两侧已存在的周次目录'
    $found = 0
    foreach ($n in @((Get-InputWeeks).Keys)) {
        $null = Get-OrCreateWeek -Number $n
        $found++
    }
    if ($found -eq 0) {
        Write-AgentLog -Message '  未发现任何 W数字公司公告 目录，等待新输入。'
    } else {
        Write-AgentLog -Message ("  已登记 {0} 个已有周次，交由稳定性判定处理。" -f $found)
    }
}

function Invoke-StabilityScan {
    if ($script:Weeks.Count -eq 0) { return }
    $now = Get-Date

    foreach ($tag in @($script:Weeks.Keys)) {
        $e = $script:Weeks[$tag]
        if ($e.Terminal) { continue }

        $paths = Get-WeekPaths -WeekTag $e.Tag
        $dirA = Join-Path (Join-Path $script:InputRoot 'A股公告') ('W{0:d2}公司公告' -f $e.Number)
        $dirHK = Join-Path (Join-Path $script:InputRoot '港股公告') ('W{0:d2}公司公告' -f $e.Number)
        $statA = Get-WeekDirInfo -Path $dirA
        $statHK = Get-WeekDirInfo -Path $dirHK

        # ---- 磁盘信号（权威）：事件丢失也不会误判
        $newest = $e.NewestSeen
        foreach ($s in @($statA, $statHK)) {
            if ($null -ne $s.Newest -and ($null -eq $newest -or $s.Newest -gt $newest)) { $newest = $s.Newest }
        }
        foreach ($t in @($e.ALastEventAt, $e.HKLastEventAt)) {
            if ($null -ne $t -and ($null -eq $newest -or $t -gt $newest)) { $newest = $t }
        }
        if ($null -ne $newest -and ($null -eq $e.NewestSeen -or $newest -gt $e.NewestSeen)) {
            $e.StableLogged = $false
        }
        $e.NewestSeen = $newest

        # ---- 事件爆发汇总（一次扫描只打一行）
        if ($e.BurstLogged -and ($e.AEvents + $e.HKEvents) -gt 1) {
            Write-AgentLog -Message ("{0} 事件汇总：A股 {1} 个 / 港股 {2} 个（A股 新增{3} 修改{4} 改名{5} 删除{6}；港股 新增{7} 修改{8} 改名{9} 删除{10}）" -f `
                $e.Tag, $e.AEvents, $e.HKEvents,
                $e.ACounts.Created, $e.ACounts.Changed, $e.ACounts.Renamed, $e.ACounts.Deleted,
                $e.HKCounts.Created, $e.HKCounts.Changed, $e.HKCounts.Renamed, $e.HKCounts.Deleted)
            $e.ACounts = @{ Created = 0; Changed = 0; Renamed = 0; Deleted = 0 }
            $e.HKCounts = @{ Created = 0; Changed = 0; Renamed = 0; Deleted = 0 }
            $e.AEvents = 0
            $e.HKEvents = 0
            $e.BurstLogged = $false
        }

        if ($null -eq $newest) { continue }

        # ---- 稳定等待（核心去抖动）
        $quietMin = ($now - $newest).TotalMinutes
        if ($quietMin -lt $StableMinutes) {
            if (-not $e.StableLogged) {
                Write-AgentLog -Level STEP -Message ("{0} 进入稳定等待：最后变化 {1}，还需 {2:N1} 分钟无变化。" -f `
                    $e.Tag, (Format-Age $newest), ($StableMinutes - $quietMin))
                $e.StableLogged = $true
            }
            continue
        }
        if (-not $e.StableLogged) {
            Write-AgentLog -Level OK -Message ("{0} 已连续 {1} 分钟无变化 → 进入正式判定。" -f $e.Tag, $StableMinutes)
            $e.StableLogged = $true
        }

        if ($e.Dispatched) { continue }

        if ($null -ne $e.DeferUntil -and $now -lt $e.DeferUntil) {
            if (-not $e.DeferLogged) {
                Write-AgentLog -Level SKIP -Message ("{0} 处于重试冷却期，{1} 后再判定。" -f `
                    $e.Tag, $e.DeferUntil.ToString('HH:mm:ss'))
                $e.DeferLogged = $true
            }
            continue
        }
        $e.DeferUntil = $null
        $e.DeferLogged = $false

        Invoke-WeekGate -Entry $e -Paths $paths -StatA $statA -StatHK $statHK -Newest $newest
    }
}

function Invoke-WeekGate {
    <#
      正式启动条件 1-8（规范 §十六）。全过才派发；
      任一不满足只记录并等待，不设额外等待间隔。
    #>
    param(
        [Parameter(Mandatory)]$Entry,
        [Parameter(Mandatory)]$Paths,
        [Parameter(Mandatory)]$StatA,
        [Parameter(Mandatory)]$StatHK,
        $Newest
    )
    $e = $Entry
    $tag = $e.Tag
    $blockers = New-Object System.Collections.ArrayList
    $lines = New-Object System.Collections.ArrayList
    $header = '{0} 正式判定（输出目录 {1}）' -f $tag, $Paths.ReportDir

    $check = {
        param([string]$Name, [bool]$Pass, [string]$Detail)
        if ($Pass) {
            [void]$lines.Add([pscustomobject]@{ Level = 'OK'; Message = ('  ✔ {0}：{1}' -f $Name, $Detail) })
        } else {
            [void]$lines.Add([pscustomobject]@{ Level = 'SKIP'; Message = ('  ✘ {0}：{1}' -f $Name, $Detail) })
            [void]$blockers.Add($Name)
        }
    }
    $flush = {
        Write-AgentLog -Level STEP -Message $header
        foreach ($l in $lines) { Write-AgentLog -Level $l.Level -Message $l.Message }
    }

    # 0. 只自动启动「最新共同周次」
    $newestWeek = Get-NewestCommonWeekNumber
    $isNewest = ($null -eq $newestWeek) -or ($e.Number -ge $newestWeek)
    if (-not $isNewest) {
        Write-AgentLog -Level STEP -Message $header
        Write-AgentLog -Level SKIP -Message ('  ✘ 为最新周次：两侧已出现更新的 W{0:d2}，本旧周次不再自动启动。' -f $newestWeek)
        Set-WeekTerminal -Entry $e -Reason ('已被更新的周次 W{0:d2} 取代' -f $newestWeek)
        Write-AgentLog -Level WARN -Message ("{0} 已标记为终态：已被 W{1:d2} 取代。" -f $tag, $newestWeek)
        Write-AgentLog -Message ("  如需补跑历史周次，请人工执行：.\地产公司公告Agent.ps1 run -Week {0}" -f $tag)
        return
    }
    $newestDetail = if ($null -eq $newestWeek) { '两侧暂无更新周次' } else { '两侧最新为 W{0:d2}' -f $newestWeek }
    & $check '为最新周次' $true $newestDetail

    # 1/2 两侧周次目录
    & $check 'A股存在对应周次目录' ($StatA.Exists)  $(if ($StatA.Exists)  { '02_待处理\A股公告\W{0:d2}公司公告' -f $e.Number }  else { '缺失' })
    & $check '港股存在对应周次目录' ($StatHK.Exists) $(if ($StatHK.Exists) { '02_待处理\港股公告\W{0:d2}公司公告' -f $e.Number } else { '缺失' })

    # 3/4 两侧非空
    & $check 'A股目录至少 1 个文件'  ($StatA.Count  -gt 0) ('实际 {0} 个文件' -f $StatA.Count)
    & $check '港股目录至少 1 个文件' ($StatHK.Count -gt 0) ('实际 {0} 个文件' -f $StatHK.Count)

    # 5 两侧稳定
    & $check ("两侧最近 {0} 分钟无变化" -f $StableMinutes) $true ('最后变化 {0}' -f (Format-Age $Newest))

    # 6 done.flag
    $doneExists = Test-Path -LiteralPath $Paths.DoneFlag
    & $check 'done.flag 不存在' (-not $doneExists) $(
        if ($doneExists) { "已存在 {0} → 本周已完成，不再启动" -f $Paths.DoneFlag } else { '未发现 done.flag' })
    if ($doneExists) {
        & $flush
        Set-WeekTerminal -Entry $e -Reason '本周已完成（done.flag 存在）'
        Write-AgentLog -Level OK -Message ("{0} 已标记为终态：本周已完成。" -f $tag)
        return
    }

    # 7 running 锁
    $running = Get-WatcherRunningInfo -Paths $Paths
    $isRunning = ($null -ne $running) -and $running.Alive
    & $check 'running 锁不存在' (-not $isRunning) $(
        if ($null -eq $running) { '无 .running 锁' }
        elseif ($running.Alive) { "PID {0} 仍在运行（起始 {1}）→ 跳过，不重复启动" -f $running.ProcessId, $running.Start }
        else { "发现失效锁（PID {0} 已不存在）→ 视为未在运行" -f $running.ProcessId })

    # 8 manual 标记 + 尝试上限
    $manualExists = Test-Path -LiteralPath $Paths.ManualLock
    & $check 'manual 标记不存在' (-not $manualExists) $(
        if ($manualExists) { "存在 {0} → 已停止自动尝试，等待人工处理" -f $Paths.ManualLock } else { '无 .manual 标记' })
    if ($manualExists) {
        & $flush
        Set-WeekTerminal -Entry $e -Reason '等待人工处理（.manual 存在）'
        Write-AgentLog -Level WARN -Message ("{0} 已标记为终态：存在人工介入标记。" -f $tag)
        return
    }

    $attempts = Get-WatcherAttemptCount -Paths $Paths
    $underCap = ($attempts -lt $MaxRuns)
    $attemptDetail = if ($underCap) { '已尝试 {0} 次' -f $attempts }
                     else { '已尝试 {0} 次，达到上限 → 停止自动尝试' -f $attempts }
    & $check ("自动尝试次数未达上限（{0}/{1}）" -f $attempts, $MaxRuns) $underCap $attemptDetail
    if (-not $underCap) {
        & $flush
        Set-WeekTerminal -Entry $e -Reason ('已达自动尝试上限（{0}/{1}）' -f $attempts, $MaxRuns)
        Write-AgentLog -Level ERROR -Message ("{0} 已标记为终态：自动尝试达上限，等待人工处理。" -f $tag)
        return
    }

    # ---------------- 判定结果 ----------------
    if ($blockers.Count -gt 0) {
        $key = ($blockers -join '|')
        $now = Get-Date
        $changed = ($key -ne $e.LastBlockKey)
        $heartbeat = ($null -eq $e.NextHeartbeatAt) -or ($now -ge $e.NextHeartbeatAt)
        if ($changed -or $heartbeat) {
            & $flush
            Write-AgentLog -Level WARN -Message ("{0} 条件未满足 → 本次不启动：{1}" -f $tag, ($blockers -join '；'))
            $e.NextHeartbeatAt = $now.AddMinutes($HeartbeatMinutes)
        }
        $e.LastBlockKey = $key
        return
    }

    & $flush
    $e.LastBlockKey = ''
    $e.NextHeartbeatAt = $null
    Start-WeeklyAgent -Entry $e -Paths $Paths
}

function Defer-Week {
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][int]$Minutes)
    $Entry.DeferUntil = (Get-Date).AddMinutes($Minutes)
    $Entry.DeferLogged = $false
    $Entry.StableLogged = $false
}

function Start-WeeklyAgent {
    <#
      派发正式 run。本函数只启动子进程，**自己不调用模型**。
      子进程自行加载 .env、自行弹状态窗口、自行写 done.flag。
    #>
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)]$Paths)
    $e = $Entry
    $tag = $e.Tag

    $argList = @(
        '-NoProfile', '-WindowStyle', 'Hidden', '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $script:ScriptPath),
        'run', '-Week', $tag
    )
    $shown = 'powershell.exe ' + ($argList -join ' ')
    Write-AgentLog -Level OK -Message ("{0} 全部条件满足 → 启动正式运行" -f $tag)
    Write-AgentLog -Level STEP -Message ("启动命令：{0}" -f $shown)

    if ($DryRun) {
        Write-AgentLog -Level WARN -Message '（DryRun）本应启动上面的命令，但**未启动**；不调用任何模型、不产生 API 请求。'
        Set-WeekTerminal -Entry $e -Reason 'DryRun（未真正启动）'
        return
    }

    # 写 .running 锁（含 PID），保证同一周次只有一个正式任务
    try {
        if (-not (Test-Path -LiteralPath $script:LockDir)) {
            New-Item -ItemType Directory -Path $script:LockDir -Force | Out-Null
        }
        $body = @("PID=$PID", "START=$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
                  "WEEK=$tag", "BY=watcher") -join "`r`n"
        Set-Content -LiteralPath $Paths.RunningLock -Value $body -Encoding UTF8
    } catch {
        Write-AgentLog -Level WARN -Message ("写入 running 锁失败：{0}" -f $_.Exception.Message)
    }

    # 累计尝试次数
    try {
        $n = (Get-WatcherAttemptCount -Paths $Paths) + 1
        Set-Content -LiteralPath $Paths.AttemptFile -Value "$n" -Encoding ASCII
        if ($n -ge $MaxRuns) {
            Set-Content -LiteralPath $Paths.ManualLock `
                -Value ("已达自动尝试上限（{0}/{1}），等待人工处理。" -f $n, $MaxRuns) -Encoding UTF8
            Write-AgentLog -Level WARN -Message ("已达自动尝试上限（{0}/{1}）→ 已写入 .manual。" -f $n, $MaxRuns)
        }
    } catch {
        Write-AgentLog -Level WARN -Message ("累计尝试次数失败：{0}" -f $_.Exception.Message)
    }

    try {
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList `
                    -WorkingDirectory $script:AgentRoot -WindowStyle Hidden -PassThru
        $e.Dispatched = $true
        $e.ChildProcess = $proc
        $e.ChildStartedAt = Get-Date
        Write-AgentLog -Level OK -Message ("已启动：PID {0}" -f $proc.Id)
    } catch {
        Write-AgentLog -Level ERROR -Message ("启动失败：{0}" -f $_.Exception.Message)
        Defer-Week -Entry $e -Minutes $FailureCooldownMinutes
    }
}

function Update-DispatchedChildren {
    foreach ($tag in @($script:Weeks.Keys)) {
        $e = $script:Weeks[$tag]
        if (-not $e.Dispatched) { continue }

        $proc = $e.ChildProcess
        $exited = $true
        try {
            if ($null -ne $proc) { $proc.Refresh(); $exited = $proc.HasExited } else { $exited = $true }
        } catch { $exited = $true }
        if (-not $exited) { continue }

        $rc = $null
        try { if ($null -ne $proc) { $rc = $proc.ExitCode } } catch { }

        $paths = Get-WeekPaths -WeekTag $tag
        Write-AgentLog -Message ("{0} 正式运行结束，返回状态：{1}" -f $tag, $(if ($null -eq $rc) { '（未知）' } else { $rc }))
        $e.Dispatched = $false
        $e.ChildProcess = $null

        if (Test-Path -LiteralPath $paths.DoneFlag) {
            Set-WeekTerminal -Entry $e -Reason '已成功完成（done.flag 存在）'
            Write-AgentLog -Level OK -Message ("{0} 已完成。" -f $tag)
            continue
        }
        if (Test-Path -LiteralPath $paths.ManualLock) {
            Set-WeekTerminal -Entry $e -Reason '等待人工处理（.manual 存在）'
            Write-AgentLog -Level ERROR -Message ("{0} 未完成且已出现 .manual → 停止自动尝试。" -f $tag)
            continue
        }

        Write-AgentLog -Level WARN -Message ("{0} 未成功完成（返回 {1}）。" -f $tag, $(if ($null -eq $rc) { '未知' } else { $rc }))
        Defer-Week -Entry $e -Minutes $FailureCooldownMinutes
    }
}

function Start-FolderWatcher {
    <#
      常驻监听 02_待处理\A股公告 与 02_待处理\港股公告。
      ★ 本函数及其调用的所有逻辑绝不调用模型、绝不读取 .env、绝不发起网络请求。
    #>
    $startTime = Get-Date
    $null = Initialize-AgentLog -Kind 'watcher'
    Set-Location -LiteralPath $script:AgentRoot

    Write-AgentLog -Level STEP -Message ("===== 监听器启动（{0}{1}）=====" -f `
        $script:AGENT_VERSION, $(if ($DryRun) { ' · DryRun' } else { '' }))

    if (-not (Enter-WatcherLock)) { return 0 }

    $watchers = @()
    $exitCode = 0
    try {
        $rootA = Join-Path $script:InputRoot 'A股公告'
        $rootHK = Join-Path $script:InputRoot '港股公告'

        foreach ($r in @($rootA, $rootHK)) {
            if (-not (Test-Path -LiteralPath $r -PathType Container)) {
                Write-AgentLog -Level ERROR -Message "监听目录不存在：$r"
                throw "监听目录不存在：$r"
            }
        }

        Write-Host ''
        Write-Host '====================================' -ForegroundColor Cyan
        Write-Host '地产公司公告 Agent 监听器已启动' -ForegroundColor Cyan
        Write-Host '监听A股：' -ForegroundColor Cyan
        Write-Host $rootA
        Write-Host ''
        Write-Host '监听港股：' -ForegroundColor Cyan
        Write-Host $rootHK
        Write-Host ''
        Write-Host '等待新周次输入……' -ForegroundColor Cyan
        Write-Host '====================================' -ForegroundColor Cyan
        Write-Host ''

        Write-AgentLog -Message ("监听A股：{0}" -f $rootA)
        Write-AgentLog -Message ("监听港股：{0}" -f $rootHK)
        Write-AgentLog -Message ("去抖动：同一周次连续 {0} 分钟无变化才算稳定；判定间隔 {1} 秒。" -f $StableMinutes, $ScanIntervalSeconds)
        if ($DryRun) { Write-AgentLog -Level WARN -Message 'DryRun 模式：只判定、只记日志，绝不启动任何 Agent 进程。' }

        # ---- FileSystemWatcher（事件驱动，不用高频轮询）
        foreach ($item in @(@{ Market = 'A'; Root = $rootA }, @{ Market = 'HK'; Root = $rootHK })) {
            $fsw = New-Object System.IO.FileSystemWatcher
            $fsw.Path = $item.Root
            $fsw.IncludeSubdirectories = $true
            $fsw.NotifyFilter = [System.IO.NotifyFilters]'FileName, DirectoryName, LastWrite, Size, CreationTime'
            $fsw.InternalBufferSize = 65536

            foreach ($evName in @('Created', 'Changed', 'Renamed', 'Deleted', 'Error')) {
                $src = '{0}|{1}' -f $item.Market, $evName
                Register-ObjectEvent -InputObject $fsw -EventName $evName -SourceIdentifier $src | Out-Null
            }
            $fsw.EnableRaisingEvents = $true
            $watchers += $fsw
            Write-AgentLog -Message ("已监听 {0} 侧：{1}（含子目录；Created/Changed/Renamed/Deleted）" -f $item.Market, $item.Root)
        }

        if (-not $NoInitialScan) { Initialize-PendingWeeks }
        Write-AgentLog -Level OK -Message '监听已就绪，等待文件变化……（Ctrl+C 退出）'

        # ---- 常驻主循环：Wait-Event 阻塞等事件，超时醒来做一次稳定判定
        while ($true) {
            $null = Wait-Event -Timeout $ScanIntervalSeconds
            foreach ($queued in @(Get-Event)) {
                try { Handle-AgentFileEvent -Event $queued }
                catch { Write-AgentLog -Level ERROR -Message ("处理文件事件出错：{0}" -f $_.Exception.Message) }
                Remove-Event -EventIdentifier $queued.EventIdentifier -ErrorAction SilentlyContinue
            }

            try { Update-DispatchedChildren }
            catch { Write-AgentLog -Level ERROR -Message ("检查子进程状态出错：{0}" -f $_.Exception.Message) }

            try { Invoke-StabilityScan }
            catch { Write-AgentLog -Level ERROR -Message ("稳定判定出错：{0}" -f $_.Exception.Message) }

            $script:Cycles++
            if ($MaxCycles -gt 0 -and $script:Cycles -ge $MaxCycles) {
                Write-AgentLog -Message ("已达到 -MaxCycles {0}，正常退出（测试模式）。" -f $MaxCycles)
                break
            }
        }
    } catch {
        Write-AgentLog -Level ERROR -Message ("监听器异常：{0}" -f $_.Exception.Message)
        $exitCode = 1
    } finally {
        foreach ($src in @('A|Created', 'A|Changed', 'A|Renamed', 'A|Deleted', 'A|Error',
                           'HK|Created', 'HK|Changed', 'HK|Renamed', 'HK|Deleted', 'HK|Error')) {
            try { Unregister-Event -SourceIdentifier $src -ErrorAction SilentlyContinue } catch { }
        }
        foreach ($fsw in $watchers) {
            try { $fsw.EnableRaisingEvents = $false; $fsw.Dispose() } catch { }
        }
        Exit-WatcherLock
        Write-AgentLog -Level STEP -Message ("===== 监听器退出（运行 {0:N1} 分钟）=====" -f ((Get-Date) - $startTime).TotalMinutes)
    }
    return $exitCode
}

# ================================================================== status
function Get-AgentProgress {
    <#
      从日志尾部解析最近一条 [STEP] [i/n] 说明，算出百分比。
      ★ 这是 **Agent 流程步骤** 的完成度，不是模型内部 token 完成度。
    #>
    param([string[]]$Lines)
    $step = 0; $total = 6; $label = '等待开始'
    foreach ($line in $Lines) {
        if ($line -match '\[STEP\]\s*\[(\d+)\/(\d+)\]\s*(.*)$') {
            $step = [int]$Matches[1]
            $total = [int]$Matches[2]
            $label = $Matches[3].Trim()
        }
    }
    $percent = 0
    if ($total -gt 0) {
        $percent = [math]::Min(100, [math]::Max(0, [math]::Round(($step / $total) * 100)))
    }
    if ($Lines -match '运行完成|已完成|done') {
        if ($step -ge $total -and $total -gt 0) { $percent = 100 }
    }
    return [pscustomobject]@{ Step = $step; Total = $total; Label = $label; Percent = $percent }
}

function Get-AgentProcesses {
    $items = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match '^(claude|node|python|pythonw|powershell|pwsh)(\.exe)?$' -and (
            $_.CommandLine -match 'agent_core' -or
            $_.CommandLine -match '地产公司公告Agent' -or
            $_.CommandLine -match 'claude\s+-p'
        )
    }
    if (-not $items) { return @() }
    $out = @()
    foreach ($p in $items) {
        $proc = Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue
        $out += [pscustomobject]@{
            Name = $p.Name
            Id = $p.ProcessId
            CPU = if ($proc -and $null -ne $proc.CPU) { [math]::Round($proc.CPU, 1) } else { 0 }
            CommandLine = $p.CommandLine
        }
    }
    return $out
}

function Get-StatusWindowProcess {
    return @(Get-Process -ErrorAction SilentlyContinue |
             Where-Object { $_.MainWindowTitle -eq $script:STATUS_TITLE })
}

function Show-AgentStatus {
    <#
      -Detached：以独立进程拉起窗口（run 开工时用；已存在则不重复弹）
      不加开关：在本进程内直接显示窗口（status 动词用）
    #>
    param([switch]$Detached)

    if (-not $Detached) {
        return Show-AgentStatusWindow
    }

    try {
        if ((Get-StatusWindowProcess).Count -gt 0) {
            Write-AgentLog -Message '状态窗口已在运行，不重复弹出。'
            return 0
        }
        if (Test-Path -LiteralPath $script:StatusLock) {
            $old = 0
            try { $old = [int]((Get-Content -LiteralPath $script:StatusLock -Raw -ErrorAction Stop).Trim()) } catch { }
            if ($old -gt 0 -and (Test-ProcessAlive -ProcessId $old)) {
                Write-AgentLog -Message '状态窗口已在运行（锁文件判据），不重复弹出。'
                return 0
            }
            Remove-Item -LiteralPath $script:StatusLock -Force -ErrorAction SilentlyContinue
        }
        if (-not (Test-Path -LiteralPath $script:LockDir)) {
            New-Item -ItemType Directory -Path $script:LockDir -Force | Out-Null
        }
        $proc = Start-Process -FilePath 'powershell.exe' `
            -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $script:ScriptPath),
                            'status', '-NoWindow') `
            -WorkingDirectory $script:AgentRoot -PassThru
        Set-Content -LiteralPath $script:StatusLock -Value $proc.Id -Encoding ASCII
        Write-AgentLog -Level OK -Message ("已弹出状态窗口（PID {0}，独立进程，不阻塞主任务）。" -f $proc.Id)
    } catch {
        # 弹窗失败绝不拖垮主任务
        Write-AgentLog -Level WARN -Message ("状态窗口启动失败（不影响主任务）：{0}" -f $_.Exception.Message)
    }
    return 0
}

function Show-AgentStatusWindow {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $script:STATUS_TITLE           # 必须与 $script:STATUS_TITLE 一致
    $form.Size = New-Object System.Drawing.Size(860, 620)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = '地产公司公告 Agent 运行状态'
    $title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 16, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(20, 15)
    $title.Size = New-Object System.Drawing.Size(560, 35)
    $form.Controls.Add($title)

    function New-Label([string]$text, [int]$y, [int]$h) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Location = New-Object System.Drawing.Point(20, $y)
        $l.Size = New-Object System.Drawing.Size(800, $h)
        $form.Controls.Add($l)
        return $l
    }

    $statusLabel  = New-Label '正在检测...' 60 28
    $weekLabel    = New-Label '当前周次：-' 92 25
    $stepLabel    = New-Label '当前进度：-' 122 25

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = New-Object System.Drawing.Point(20, 152)
    $progress.Size = New-Object System.Drawing.Size(800, 28)
    $progress.Minimum = 0
    $progress.Maximum = 100
    $form.Controls.Add($progress)

    $percentLabel = New-Label '0%' 184 25
    $processLabel = New-Label '运行进程：-' 216 60
    $pidLabel     = New-Label 'PID / CPU：-' 278 25
    $doneLabel    = New-Label 'done：-' 306 25
    $outLabel     = New-Label ('输出目录：{0}' -f $script:OutRoot) 334 25

    $logBox = New-Object System.Windows.Forms.TextBox
    $logBox.Location = New-Object System.Drawing.Point(20, 366)
    $logBox.Size = New-Object System.Drawing.Size(800, 160)
    $logBox.Multiline = $true
    $logBox.ScrollBars = 'Vertical'
    $logBox.ReadOnly = $true
    $logBox.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($logBox)

    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Text = '立即刷新'
    $refreshButton.Location = New-Object System.Drawing.Point(20, 536)
    $refreshButton.Size = New-Object System.Drawing.Size(110, 32)
    $form.Controls.Add($refreshButton)

    $openOutputButton = New-Object System.Windows.Forms.Button
    $openOutputButton.Text = '打开输出目录'
    $openOutputButton.Location = New-Object System.Drawing.Point(145, 536)
    $openOutputButton.Size = New-Object System.Drawing.Size(130, 32)
    $form.Controls.Add($openOutputButton)

    $refresh = {
        # 最新日志
        $latestLog = $null
        if (Test-Path -LiteralPath $script:LogDir) {
            $latestLog = Get-ChildItem -LiteralPath $script:LogDir -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
        }
        $lines = @()
        if ($latestLog) {
            try { $lines = @(Get-Content -LiteralPath $latestLog.FullName -Tail 120 -ErrorAction Stop) }
            catch { $lines = @("读取日志失败：$($_.Exception.Message)") }
        }

        # 最新周次
        $week = '-'
        if (Test-Path -LiteralPath $script:OutRoot) {
            $d = Get-ChildItem -LiteralPath $script:OutRoot -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d{4}-W\d{2}$' } |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($d) { $week = $d.Name }
        }
        $weekLabel.Text = "当前周次：$week"

        # 进度（Agent 流程步骤完成度，非模型 token 完成度）
        $prog = Get-AgentProgress -Lines $lines
        $progress.Value = [int]$prog.Percent
        $percentLabel.Text = "$($prog.Percent)%"
        $stepLabel.Text = "当前进度：[$($prog.Step)/$($prog.Total)] $($prog.Label)"

        # 进程
        $procs = @(Get-AgentProcesses)
        if ($procs.Count -gt 0) {
            $statusLabel.Text = '状态：正在运行'
            $processLabel.Text = '运行进程：' + (($procs | ForEach-Object { "$($_.Name)  PID=$($_.Id)" }) -join '    ')
            $pidLabel.Text = 'PID / CPU：' + (($procs | ForEach-Object { "$($_.Id) / $($_.CPU)s" }) -join '    ')
        } else {
            if ($latestLog -and $latestLog.LastWriteTime -gt (Get-Date).AddMinutes(-3)) {
                $statusLabel.Text = '状态：最近有活动，当前未检测到 Agent 主进程'
            } else {
                $statusLabel.Text = '状态：空闲 / 等待新一周输入'
            }
            $processLabel.Text = '运行进程：未检测到 claude / python 活动进程'
            $pidLabel.Text = 'PID / CPU：-'
        }

        # done 状态
        $doneText = 'done：未完成'
        if ($week -ne '-') {
            $df = Join-Path (Join-Path $script:OutRoot $week) 'done.flag'
            if (Test-Path -LiteralPath $df) {
                $doneText = "done：已完成（$((Get-Item -LiteralPath $df).LastWriteTime)）"
            }
        }
        $doneLabel.Text = $doneText
        $outLabel.Text = '输出目录：' + $(if ($week -ne '-') { Join-Path $script:OutRoot $week } else { $script:OutRoot })

        if ($latestLog) {
            $logBox.Text = "最新日志：$($latestLog.Name)    更新：$($latestLog.LastWriteTime)`r`n" +
                           (($lines | Select-Object -Last 35) -join "`r`n")
            $logBox.SelectionStart = $logBox.Text.Length
            $logBox.ScrollToCaret()
        } else {
            $logBox.Text = '尚未找到日志文件。'
        }
    }

    $refreshButton.Add_Click($refresh)
    $openOutputButton.Add_Click({
        $target = $script:OutRoot
        if (Test-Path -LiteralPath $target) { Start-Process explorer.exe $target }
    })

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 2000
    $timer.Add_Tick($refresh)
    $timer.Start()

    $form.Add_Shown($refresh)
    [void]$form.ShowDialog()
    return 0
}

# ================================================================== 计划任务
function Install-AgentWatcher {
    <#
      注册登录自启动的监听计划任务。
      不写死用户名、不写死绝对目录：一律用当前脚本的真实路径。
    #>
    param([switch]$Confirmed)

    $self = $script:ScriptPath
    $taskCmd = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" watch' -f $self

    Write-Host ''
    Write-Host '即将注册 Windows 计划任务：' -ForegroundColor Cyan
    Write-Host ("  名称     : {0}" -f $script:TASK_NAME)
    Write-Host ("  触发     : 当前用户登录时自启动")
    Write-Host ("  命令     : powershell.exe {0}" -f $taskCmd)
    Write-Host ("  工作目录 : {0}" -f $script:AgentRoot)
    Write-Host ''

    # 旧任务只提示，绝不擅自删除（规范 §十九）。
    # 旧版实际注册过两个任务：周任务 + 登录常驻监听。两个都要检查——
    # 登录监听那个如果还是 Ready，会和本任务同时跑，造成重复触发。
    foreach ($oldName in $script:OLD_TASK_NAMES) {
        $oldTask = Get-ScheduledTask -TaskName $oldName -ErrorAction SilentlyContinue
        if (-not $oldTask) { continue }
        Write-Host ("提示：检测到旧计划任务 {0}（状态 {1}）。" -f $oldName, $oldTask.State) -ForegroundColor Yellow
        if ($oldTask.State -ne 'Disabled') {
            Write-Host '      ⚠ 它仍是启用状态，可能与本任务重复触发。建议先禁用：' -ForegroundColor Yellow
            Write-Host ("      Disable-ScheduledTask -TaskName {0}" -f $oldName) -ForegroundColor DarkGray
        }
        Write-Host '      本命令不会删除它。如需删除，请自行确认后执行：' -ForegroundColor Yellow
        Write-Host ("      Unregister-ScheduledTask -TaskName {0} -Confirm:`$false" -f $oldName) -ForegroundColor DarkGray
        Write-Host ''
    }

    if (-not $Confirmed) {
        # 无标准输入时（提权窗口拉起的隐藏进程）Read-Host 会永久阻塞，
        # 看起来像「命令卡死」。先探测 stdin 是否可用，不可用就直接报错退出，
        # 而不是无声地挂住。
        $stdinUsable = $true
        try { $null = [Console]::In } catch { $stdinUsable = $false }
        if (-not $stdinUsable) {
            Write-Host '当前环境没有可用的标准输入，无法交互确认。' -ForegroundColor Red
            Write-Host '请改用：地产公司公告Agent.ps1 install-watcher -Yes' -ForegroundColor Yellow
            return 1
        }
        $ans = Read-Host '确认注册？输入 yes 继续'
        if ($ans -ne 'yes') {
            Write-Host '已取消，未做任何改动。' -ForegroundColor Yellow
            return 1
        }
    }

    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskCmd `
        -WorkingDirectory $script:AgentRoot
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
        -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 0)
    $desc = '地产公司公告 Agent：登录后常驻监听 02_待处理\A股公告 与 02_待处理\港股公告；' +
            '连续 5 分钟无变化且两侧齐全才启动正式运行。监听过程不调用任何大模型、不产生 API 请求。'

    try {
        Register-ScheduledTask -TaskName $script:TASK_NAME -Action $action -Trigger $trigger `
            -Settings $settings -Description $desc -Force | Out-Null
        Write-Host ''
        Write-Host ("已注册计划任务：{0}" -f $script:TASK_NAME) -ForegroundColor Green
        Write-Host ("查看：Get-ScheduledTask -TaskName {0} | Get-ScheduledTaskInfo" -f $script:TASK_NAME)
        Write-Host '卸载：地产公司公告Agent.ps1 uninstall-watcher'
        return 0
    } catch {
        Write-Host ("注册失败：{0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-Host '（若提示权限不足，请用「以管理员身份运行」的 PowerShell 重试）' -ForegroundColor Yellow
        return 1
    }
}

function Uninstall-AgentWatcher {
    <#
      删除监听计划任务。★ 删除前必须二次确认。
    #>
    param([switch]$Confirmed)

    $existing = Get-ScheduledTask -TaskName $script:TASK_NAME -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Host ("未发现计划任务 {0}，无需删除。" -f $script:TASK_NAME) -ForegroundColor Yellow
        return 0
    }

    Write-Host ''
    Write-Host '即将删除 Windows 计划任务：' -ForegroundColor Cyan
    Write-Host ("  名称 : {0}" -f $script:TASK_NAME)
    Write-Host ("  状态 : {0}" -f $existing.State)
    Write-Host ''
    Write-Host '（只删除计划任务，不删除任何文件、不触碰待处理公告与输出。）' -ForegroundColor DarkGray
    Write-Host ''

    if (-not $Confirmed) {
        $ans = Read-Host '确认删除？输入 yes 继续'
        if ($ans -ne 'yes') {
            Write-Host '已取消，未做任何改动。' -ForegroundColor Yellow
            return 1
        }
    }

    try {
        Unregister-ScheduledTask -TaskName $script:TASK_NAME -Confirm:$false -ErrorAction Stop
        Write-Host ("已删除计划任务：{0}" -f $script:TASK_NAME) -ForegroundColor Green
        return 0
    } catch {
        Write-Host ("删除失败：{0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-Host '（若提示权限不足，请用「以管理员身份运行」的 PowerShell 重试）' -ForegroundColor Yellow
        return 1
    }
}

# ================================================================== build-release
function Test-ReleaseForSensitive {
    <#
      发行版敏感信息扫描（规范 §二十六）。返回 @{ Clean; Hits }。
      *** 命中即阻断，绝不生成 ZIP。 ***

      判据分两类，刻意分开：
        · 文件名 —— 真正的 .env / 个人配置.json 文件本身出现在包里，才是真泄漏。
          文档里写「复制 .env.example 为 .env」是说明书，不是泄漏，不能误杀。
        · 文件内容 —— 私人绝对路径、裸用户号、真的被填了值的 API Key、各种 token 前缀。

      下面这些判据字面量一律用字符串拼接拼出来，**不让它们以完整形态出现在本文件源码里**，
      否则扫描会命中自己的源码（同文自指）。例如 '152' + '58' 就拼不出连续的五个数字。
    #>
    param([Parameter(Mandatory)][string]$Root)

    $userNo   = '152' + '58'
    $userPath = 'C:\Users\' + $userNo
    $intern   = 'Desktop\' + '实习'
    $tokenKey = 'ANTHROPIC_' + 'AUTH_TOKEN'

    $hits = New-Object System.Collections.ArrayList
    $textExt = @('.txt', '.md', '.json', '.ps1', '.py', '.csv', '.yml', '.yaml',
                 '.cfg', '.ini', '.xml', '.bat', '.cmd', '.js', '.html', '.example')

    foreach ($f in @(Get-ChildItem -LiteralPath $Root -File -Recurse -Force -ErrorAction SilentlyContinue)) {
        if ($f.FullName -match '\\\.git\\') { continue }

        # ---- 文件名判据 ----
        if ($f.Name -like '*.env' -or $f.Name -eq '.env') {
            [void]$hits.Add(("{0} :: 文件名是 .env（真实密钥文件不得进入发行版）" -f $f.FullName))
        }
        if ($f.Name -eq '个人配置.json') {
            [void]$hits.Add(("{0} :: 文件名是 个人配置.json（个人数据不得进入发行版）" -f $f.FullName))
        }

        # ---- 内容判据 ----
        if ($textExt -notcontains $f.Extension.ToLower()) { continue }
        if ($f.Length -gt 8MB) { continue }
        try {
            $content = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 -ErrorAction Stop
        } catch { continue }
        if ([string]::IsNullOrEmpty($content)) { continue }

        if ($content.Contains($userPath)) {
            [void]$hits.Add(("{0} :: 内容含私人绝对路径 {1}" -f $f.FullName, $userPath))
        }
        if ($content.Contains($userNo)) {
            [void]$hits.Add(("{0} :: 内容含裸用户号" -f $f.FullName))
        }
        if ($content.Contains($intern)) {
            [void]$hits.Add(("{0} :: 内容含历史实习目录 {1}" -f $f.FullName, $intern))
        }
        foreach ($pat in @('sk-[A-Za-z0-9]{10,}', 'ghp_[A-Za-z0-9]{10,}', 'github_pat_[A-Za-z0-9_]{10,}')) {
            if ($content -match $pat) {
                [void]$hits.Add(("{0} :: 内容含令牌前缀 {1}" -f $f.FullName, $pat.Split('[')[0]))
            }
        }

        # API Key 逐行看：只在「确实填了非空值」时才判命中。
        # 空值（如 .env.example 的 ANTHROPIC_AUTH_TOKEN=）是合规的，不算命中。
        foreach ($line in ($content -split "`r?`n")) {
            $t = $line.Trim()
            if ($t.StartsWith('#')) { continue }
            $i = $t.IndexOf('=')
            if ($i -lt 1) { continue }
            if ($t.Substring(0, $i).Trim() -ne $tokenKey) { continue }
            $val = $t.Substring($i + 1).Trim().Trim('"').Trim("'")
            if ($val.Length -gt 0) {
                [void]$hits.Add(("{0} :: {1} 已被填入真实值（必须留空）" -f $f.FullName, $tokenKey))
            }
        }
    }

    return [pscustomobject]@{ Clean = ($hits.Count -eq 0); Hits = @($hits) }
}

function Build-AgentRelease {
    <#
      生成可分享发行包（规范 §二十五 / §二十八）。
      只读取、只复制；绝不修改个人版的 .env / 个人配置.json / 工作区 / 输出 /
      日志 / 待处理 / 历史资料 / 模板原件。
    #>
    # 计数器放脚本作用域：嵌套函数读不到本函数的局部变量，显式共享更稳。
    $script:RelTotal = 22
    $script:RelStep = 0
    function Step([string]$text) {
        $script:RelStep++
        Write-AgentLog -Level STEP -Message ("[{0}/{1}] {2}" -f $script:RelStep, $script:RelTotal, $text)
    }

    Write-AgentLog -Level STEP -Message '===== 生成发行版 ====='
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $stage = Join-Path $script:WorkRoot ("_release_stage_$stamp")
    $releaseName = '地产公司公告Agent'
    $target = Join-Path $stage $releaseName

    Step '检查个人版是否具备必需文件'
    foreach ($rel in @('.env.example', '.gitignore', 'README.md', 'requirements.txt',
                       '地产公司公告Agent.ps1', 'internal\agent_core.py')) {
        $p = Join-Path $script:AgentRoot $rel
        if (-not (Test-Path -LiteralPath $p)) {
            Write-AgentLog -Level ERROR -Message "  缺少必需文件：$rel → 终止（不生成 ZIP）"
            return 1
        }
    }
    Write-AgentLog -Level OK -Message '  必需文件齐全'

    Step '准备暂存目录'
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
    New-Item -ItemType Directory -Path $target -Force | Out-Null

    Step '复制主入口脚本'
    Copy-Item -LiteralPath (Join-Path $script:AgentRoot '地产公司公告Agent.ps1') -Destination $target

    Step '复制 internal\（内部 Python 模块）'
    Copy-Item -LiteralPath $script:InternalDir -Destination $target -Recurse
    $pycache = Join-Path $target 'internal\__pycache__'
    if (Test-Path -LiteralPath $pycache) { Remove-Item -LiteralPath $pycache -Recurse -Force }

    Step '复制 00_配置\（不含 个人配置.json）'
    $cfgTarget = Join-Path $target '00_配置'
    New-Item -ItemType Directory -Path $cfgTarget -Force | Out-Null
    foreach ($f in @('Agent配置.json', '规则.md', '任务卡.txt')) {
        $p = Join-Path $script:ConfigDir $f
        if (Test-Path -LiteralPath $p) { Copy-Item -LiteralPath $p -Destination $cfgTarget }
    }
    Write-AgentLog -Message '  已排除：个人配置.json'

    Step '复制 01_模板\ 原件'
    if (Test-Path -LiteralPath $script:TemplateDir) {
        $tplTarget = Join-Path $target '01_模板'
        New-Item -ItemType Directory -Path $tplTarget -Force | Out-Null
        foreach ($f in @(Get-ChildItem -LiteralPath $script:TemplateDir -File -ErrorAction SilentlyContinue)) {
            Copy-Item -LiteralPath $f.FullName -Destination $tplTarget
        }
    }

    Step '复制 requirements.txt'
    Copy-Item -LiteralPath (Join-Path $script:AgentRoot 'requirements.txt') -Destination $target

    Step '复制 README.md'
    Copy-Item -LiteralPath (Join-Path $script:AgentRoot 'README.md') -Destination $target

    Step '写入 .env.example（只含空值/示例值，绝不含真实 Key）'
    $envEx = @(
        '# 复制本文件为 .env，然后填写自己的 API Key',
        '# 不要把填写好的 .env 提交到 Git，也不要分享给别人',
        'ANTHROPIC_BASE_URL=https://api.deepseek.com/anthropic',
        'ANTHROPIC_AUTH_TOKEN=',
        'ANTHROPIC_MODEL='
    ) -join "`r`n"
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Join-Path $target '.env.example'), $envEx, $utf8Bom)

    Step '复制 .gitignore'
    Copy-Item -LiteralPath (Join-Path $script:AgentRoot '.gitignore') -Destination $target

    Step '建立空目录骨架'
    foreach ($d in @('02_待处理\A股公告', '02_待处理\港股公告', '02_待处理\研报', '02_待处理\新闻资讯',
                     '05_输出', '04_工作区\locks', '日志', '发行版')) {
        New-Item -ItemType Directory -Path (Join-Path $target $d) -Force | Out-Null
    }

    Step '放置目录占位说明'
    $readme = '本目录用于存放输入与产出。发行版不含任何个人数据。'
    foreach ($d in @('02_待处理', '05_输出', '04_工作区', '日志')) {
        $p = Join-Path $target ("{0}\说明.txt" -f $d)
        [System.IO.File]::WriteAllText($p, $readme, $utf8Bom)
    }

    Step '确认不含 03_个人资料\ / 个人配置.json / 旧版归档 / 开发工具'
    foreach ($p in @((Join-Path $target '03_个人资料'),
                     (Join-Path $target '00_配置\个人配置.json'),
                     (Join-Path $target '兼容旧版'),
                     (Join-Path $target '开发工具'),
                     (Join-Path $target '.git'))) {
        if (Test-Path -LiteralPath $p) {
            Write-AgentLog -Level ERROR -Message "  发行版中不应存在：$p → 终止"
            return 1
        }
    }
    Write-AgentLog -Level OK -Message '  已确认不含个人资料 / 个人配置 / 兼容旧版 / 开发工具 / .git'

    Step '确认不含 .env（只允许 .env.example）'
    if (Test-Path -LiteralPath (Join-Path $target '.env')) {
        Write-AgentLog -Level ERROR -Message '  发行版中不应存在 .env → 终止'
        return 1
    }
    Write-AgentLog -Level OK -Message '  已确认只含 .env.example'

    Step '排除历史资料与旧脚本目录'
    # 发行版采用**白名单**复制：这些目录本来就不会被复制进来。
    # 这里再显式清一遍，是为了防止将来有人改成整目录复制时把个人/旧版内容带出去。
    foreach ($d in @('00_规则', '03_参考资料', '提示词', '脚本', '03_个人资料',
                     '兼容旧版', '开发工具')) {
        $p = Join-Path $target $d
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
    }
    Write-AgentLog -Level OK -Message '  已排除：00_规则 / 03_参考资料 / 提示词 / 脚本 / 个人资料 / 兼容旧版 / 开发工具'

    Step '排除输出、日志、工作区中的个人内容'
    foreach ($d in @('05_输出', '日志', '04_工作区')) {
        $p = Join-Path $target $d
        if (-not (Test-Path -LiteralPath $p)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $p -File -Recurse -ErrorAction SilentlyContinue)) {
            if ($f.Name -ne '说明.txt') { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue }
        }
    }
    Write-AgentLog -Level OK -Message '  已清空个人产出'

    Step '排除 发行版\ 中的历史 ZIP'
    $relDir = Join-Path $target '发行版'
    if (Test-Path -LiteralPath $relDir) {
        foreach ($f in @(Get-ChildItem -LiteralPath $relDir -File -ErrorAction SilentlyContinue)) {
            Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
        }
    }

    Step '排除 __pycache__ / *.pyc / *.log'
    foreach ($f in @(Get-ChildItem -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue |
                     Where-Object { $_.Name -eq '__pycache__' -or $_.Extension -in @('.pyc', '.log') })) {
        Remove-Item -LiteralPath $f.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }

    Step '写入版本信息（便于对方确认拿到的是哪一版）'
    $verTxt = @(
        "地产公司公告 Agent"
        "版本：$script:AGENT_VERSION"
        "打包时间：$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
        ""
        "本包不含 .env、个人配置.json、个人资料、兼容旧版、开发工具、输出、日志、Git 信息。"
        "使用前请先阅读 README.md。"
    ) -join "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $target 'VERSION.txt'), $verTxt, $utf8Bom)

    Step '排除 Git 信息'
    foreach ($d in @('.git', '.github', '.gitattributes', '.gitmodules')) {
        $p = Join-Path $target $d
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Write-AgentLog -Level OK -Message '  已排除 Git 信息'

    $tree = @(Get-ChildItem -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue)
    Write-AgentLog -Message ("  暂存目录共 {0} 个条目" -f $tree.Count)

    Step '扫描敏感信息（私人路径 / 密钥 / 个人配置 / .env）'
    $scan = Test-ReleaseForSensitive -Root $target
    if (-not $scan.Clean) {
        Write-Host ''
        Write-Host '[SECURITY BLOCK]' -ForegroundColor Red
        Write-Host '发行版检测到疑似私人或密钥信息。' -ForegroundColor Red
        Write-Host '不要生成ZIP。' -ForegroundColor Red
        Write-Host ''
        foreach ($h in $scan.Hits) { Write-AgentLog -Level ERROR -Message ("  {0}" -f $h) }
        Write-AgentLog -Level ERROR -Message '已终止，未生成 ZIP。请修正后重试。'
        Write-AgentLog -Level WARN  -Message ("暂存目录保留以供检查：{0}" -f $stage)
        return 1
    }
    Write-AgentLog -Level OK -Message '  扫描通过：未发现私人或密钥信息'

    Step '生成 ZIP'
    if (-not (Test-Path -LiteralPath $script:ReleaseDir)) {
        New-Item -ItemType Directory -Path $script:ReleaseDir -Force | Out-Null
    }
    $zipPath = Join-Path $script:ReleaseDir ("{0}.zip" -f $releaseName)
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    try {
        Compress-Archive -Path (Join-Path $target '*') -DestinationPath $zipPath -Force -ErrorAction Stop
    } catch {
        Write-AgentLog -Level ERROR -Message ("打包失败：{0}" -f $_.Exception.Message)
        return 1
    }

    # 下面是打包后的收尾，不再占用 22 步编号
    # （§二十五 的第 20/21/22 步固定是「排除Git信息 / 扫描敏感信息 / 生成ZIP」）
    Write-AgentLog -Message '—— 清理暂存目录'
    try { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction Stop } catch { }

    $size = [math]::Round((Get-Item -LiteralPath $zipPath).Length / 1KB, 1)
    Write-AgentLog -Level OK -Message ("发行包已生成：{0}（{1} KB）" -f $zipPath, $size)
    Write-Host ''
    Write-Host "发行包：$zipPath" -ForegroundColor Green
    Write-Host '说明：包内不含 .env、个人配置.json、个人资料、输出、日志、Git 信息。' -ForegroundColor Green
    return 0
}

# ================================================================== 帮助与分发
function Show-AgentHelp {
    Write-Host ''
    Write-Host '地产公司公告 Agent' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  init               建立目录、检查依赖与 .env'
    Write-Host '  run                运行本周（自动识别最新共同周次）'
    Write-Host '  run -Week 2026-W38 运行指定周次'
    Write-Host '  run -DryRun        只做校验，不调用模型'
    Write-Host '  watch              实时监听待处理目录（常驻）'
    Write-Host '  status             打开状态窗口'
    Write-Host '  install-watcher    注册登录自启动计划任务（需确认）'
    Write-Host '  install-watcher -Yes  同上，但跳过交互确认（需在提权窗口里执行）'
    Write-Host '  uninstall-watcher  删除计划任务（需二次确认）'
    Write-Host '  build-release      生成可分享发行包'
    Write-Host ''
    Write-Host ("  工作区：{0}" -f $script:AgentRoot) -ForegroundColor DarkGray
    Write-Host ''
    return 0
}

Set-Location -LiteralPath $script:AgentRoot

switch ($Command) {
    'init'              { exit (Initialize-Agent) }
    'run'               { exit (Start-WeeklyRun -IsDryRun:$DryRun) }
    'watch'             { exit (Start-FolderWatcher) }
    'status'            { exit (Show-AgentStatus) }
    'install-watcher'   { exit (Install-AgentWatcher -Confirmed:$Yes) }
    'uninstall-watcher' { exit (Uninstall-AgentWatcher -Confirmed:$Yes) }
    'build-release'     { exit (Build-AgentRelease) }
    default             { exit (Show-AgentHelp) }
}
