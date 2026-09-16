#Requires -Version 5.1
<#
================================================================================
 检查并触发周报.ps1 — 地产周报 Agent 自动检查触发器
================================================================================
 版本：v1.0 ／ 建立日期：2026-09-16

 作用
   被 Windows 计划任务每 15 分钟拉起一次，只做**纯本地**的轻量判断：
   本周的公司公告是否已经投放完毕、值得启动 Agent。
   **只有全部条件满足**，才启动 脚本\run_weekly.ps1；否则**立即结束本次检查**。

 ★★ 安全承诺（本脚本的硬约束，改动时不得破坏）★★
   · 检查过程**绝不调用任何大模型**，**绝不产生 DeepSeek / Anthropic API 请求**。
   · 本脚本**从不读取 .env**——进程里根本没有 API Key 可用。
     Key 只由 run_weekly.ps1 在被启动的那一刻自行加载，检查逻辑碰不到它。
   · 本脚本**不发起任何网络请求**：无 Invoke-WebRequest / Invoke-RestMethod /
     curl / wget / Start-BitsTransfer，只读本地文件系统。
   · 条件不满足 / 已跑过 / 正在跑 / 已叫停 —— 一律**立即退出，不启动 Agent**。

 触发条件（全部满足才启动，任一不满足即退出并说明原因）
   1. A股 / 港股 两侧都存在输入目录
   2. 两侧的「最新周次」是**同一个**周次
   3. 该周次的两侧目录都**非空**
   4. 两侧所有文件**最近 N 分钟都没有变化**（默认 5 分钟；正在复制 PDF 时不触发）
   5. 该周次**尚未完成**（无 05_输出\YYYY-Wxx\done.flag）
   6. 该周次**未被人工叫停**（无 04_工作区\locks\YYYY-Wxx.manual）
   7. 该周次**没有正在运行**（无有效 PID 的 .running 锁）
   8. 该周次的自动尝试次数**未达上限**（默认 2 次）
   9. 当前时间在允许窗口内（周五，且不晚于 20:00）

 用法
   .\脚本\检查并触发周报.ps1                        # 正常检查（计划任务调用）
   .\脚本\检查并触发周报.ps1 -DryRun                # 只判断、只打印，什么都不动
   .\脚本\检查并触发周报.ps1 -Week 2026-W37 -AnyDay -DryRun   # 指定周次测试
   .\脚本\检查并触发周报.ps1 -RegisterTask          # 注册计划任务（只做一次，需人工确认）

 计划任务（由 -RegisterTask 创建；本脚本**不会**自动注册）
   名称    RealEstateWeeklyWatcher
   计划    每周五 09:00 开始，每 15 分钟一次，持续 11 小时（最后一次 20:00）
   命令    powershell.exe -NoProfile -ExecutionPolicy Bypass -File <本脚本>
   工作目录 C:\Users\15258\Desktop\地产周报Agent

 日志
   日志\watcher_YYYYMMDD.log （按天追加，含每一次检查的逐项结论）
================================================================================
#>
[CmdletBinding()]
param(
    # 工作区根目录
    [string]$BaseDir = 'C:\Users\15258\Desktop\地产周报Agent',

    # 强制指定周次（形如 2026-W38）。不传则按「最近的、已结束的周五」推算。
    # 人工补跑历史周次时用；计划任务不要传。
    [string]$Week,

    # 文件稳定窗口（分钟）。两侧所有文件在这么多分钟内都没有变化才算稳定。
    [int]$StableMinutes = 5,

    # 同一周次最多自动尝试几次（每次 = 启动一次 run_weekly.ps1）
    [int]$MaxRuns = 2,

    # 当日触发截止时间（含）。晚于该时刻不再触发。
    [datetime]$CutoffTime = '20:00',

    # 只判断、只打印，不写任何状态文件、不启动 run_weekly（测试用）
    [switch]$DryRun,

    # 允许非周五运行（人工补跑 / 测试用；计划任务不要传）
    [switch]$AnyDay,

    # 忽略 done.flag（仅用于测试「已完成则不再触发」这条规则的反面场景）
    [switch]$IgnoreDone,

    # 强制忽略「本周目标周次」一致性校验（仅测试用）
    [switch]$IgnoreTargetWeek,

    # 注册 Windows 计划任务后退出（不会顺带执行检查）
    [switch]$RegisterTask
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ 常量
$WATCHER_VERSION = 'v1.0'
$RUN_WEEKLY      = '脚本\run_weekly.ps1'
$LOCK_REL        = '04_工作区\locks'
$LOG_REL         = '日志'
$TASK_NAME       = 'RealEstateWeeklyWatcher'

# 计划任务参数（与 -RegisterTask 使用的一致，便于核对）
$TASK_DAY        = 'Friday'
$TASK_AT         = '09:00'
$TASK_EVERY_MIN  = 15
$TASK_DURATION_H = 11

# ------------------------------------------------------------------ 编码
# PS 5.1 默认 $OutputEncoding 是 ASCII，会把中文输出毁掉。
$OutputEncoding = New-Object System.Text.UTF8Encoding $false
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

# ------------------------------------------------------------------ 日志
$script:LogPath = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP', 'SKIP')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        'SKIP'  { Write-Host $line -ForegroundColor DarkGray }
        default { Write-Host $line }
    }
    if ($script:LogPath) {
        try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 } catch { }
    }
}

# ------------------------------------------------------------------ ISO 周
# 与 run_weekly.ps1 同一套算法，保证两边算出的周次完全一致。
function Get-IsoWeek {
    param([Parameter(Mandatory)][datetime]$Date)
    $d = $Date.Date
    $isoDow = [int]$d.DayOfWeek
    if ($isoDow -eq 0) { $isoDow = 7 }
    $thursday = $d.AddDays(4 - $isoDow)
    $isoYear  = $thursday.Year
    $jan4     = Get-Date -Year $isoYear -Month 1 -Day 4
    $j4dow    = [int]$jan4.DayOfWeek
    if ($j4dow -eq 0) { $j4dow = 7 }
    $week1Mon = $jan4.AddDays(-($j4dow - 1))
    $week     = [int][math]::Floor(($thursday - $week1Mon).TotalDays / 7) + 1
    return [pscustomobject]@{ IsoYear = $isoYear; IsoWeek = $week }
}

# 由 ISO 周反推该周的周五（= run_weekly.ps1 的 -EndDate）
function Get-IsoWeekFriday {
    param([Parameter(Mandatory)][int]$IsoYear, [Parameter(Mandatory)][int]$IsoWeek)
    $jan4     = Get-Date -Year $IsoYear -Month 1 -Day 4
    $j4dow    = [int]$jan4.DayOfWeek
    if ($j4dow -eq 0) { $j4dow = 7 }
    $week1Mon = $jan4.AddDays(-($j4dow - 1))
    return $week1Mon.AddDays(($IsoWeek - 1) * 7 + 4)      # 周一 + 4 天 = 周五
}

# 本周目标周次：取「最近的、已结束的周五」（今天是周五就用今天），
# 与 run_weekly.ps1 的推算口径完全一致。
function Get-TargetWeek {
    param([string]$Forced)
    if ($Forced) {
        $m = [regex]::Match($Forced, '^(\d{4})-W(\d{1,2})$')
        if (-not $m.Success) { throw "周次格式不对：$Forced（期望形如 2026-W38）" }
        $y = [int]$m.Groups[1].Value
        $n = [int]$m.Groups[2].Value
        return [pscustomobject]@{
            Tag = '{0}-W{1:d2}' -f $y, $n
            IsoYear = $y; IsoWeek = $n; WeekTag = 'W{0:d2}' -f $n
            Friday = (Get-IsoWeekFriday -IsoYear $y -IsoWeek $n)
            Derived = $false
        }
    }
    $today = (Get-Date).Date
    $dow   = [int]$today.DayOfWeek
    $back  = ($dow - 5 + 7) % 7
    $friday = if ($back -eq 0) { $today } else { $today.AddDays(-$back) }
    $iso   = Get-IsoWeek -Date $friday
    return [pscustomobject]@{
        Tag = '{0}-W{1:d2}' -f $iso.IsoYear, $iso.IsoWeek
        IsoYear = $iso.IsoYear; IsoWeek = $iso.IsoWeek
        WeekTag = 'W{0:d2}' -f $iso.IsoWeek
        Friday = $friday
        Derived = $true
    }
}

# ------------------------------------------------------------------ 输入目录
# 02_待处理\<市场>\Wnn公司公告  —— 只认这个命名，其他一律忽略。
function Get-WeekDirs {
    param([Parameter(Mandatory)][string]$Parent)
    $out = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $Parent -PathType Container)) { return @() }
    foreach ($d in @(Get-ChildItem -LiteralPath $Parent -Directory -ErrorAction SilentlyContinue)) {
        $m = [regex]::Match($d.Name, '^W(\d{1,2})公司公告$')
        if ($m.Success) {
            $n = [int]$m.Groups[1].Value
            [void]$out.Add([pscustomobject]@{
                Number = $n
                Tag    = 'W{0:d2}' -f $n
                Name   = $d.Name
                Path   = $d.FullName
            })
        }
    }
    return @($out | Sort-Object Number)
}

# 某个目录下的全部文件，以及其中最新的写时间
function Get-DirFileStat {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return [pscustomobject]@{ Exists = $false; Count = 0; Newest = $null }
    }
    $files = @(Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue)
    $newest = $null
    foreach ($f in $files) {
        if ($null -eq $newest -or $f.LastWriteTime -gt $newest) { $newest = $f.LastWriteTime }
    }
    return [pscustomobject]@{ Exists = $true; Count = $files.Count; Newest = $newest }
}

function Format-Age {
    param($Newest)
    if ($null -eq $Newest) { return '（无文件）' }
    $min = ((Get-Date) - $Newest).TotalMinutes
    if ($min -lt 0)  { return ('{0:N1} 分钟前（写入时间在未来，时钟可能不准）' -f $min) }
    if ($min -lt 60) { return ('{0:N1} 分钟前' -f $min) }
    return ('{0:N1} 小时前' -f ($min / 60))
}

# ------------------------------------------------------------------ 状态文件
# 锁目录（04_工作区\locks）只在**真正要写状态**时创建；DryRun 与"未触发"路径都不碰它。
function New-LockDir {
    $d = Join-Path $BaseDir $LOCK_REL
    if (-not (Test-Path -LiteralPath $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
    return $d
}

function Get-WeekPaths {
    param([Parameter(Mandatory)][string]$WeekTag)     # 形如 2026-W38
    $lock = Join-Path $BaseDir $LOCK_REL
    return [pscustomobject]@{
        ReportDir   = Join-Path (Join-Path $BaseDir '05_输出') $WeekTag
        DoneFlag    = Join-Path (Join-Path (Join-Path $BaseDir '05_输出') $WeekTag) 'done.flag'
        RunningLock = Join-Path $lock ("$WeekTag.running")
        ManualLock  = Join-Path $lock ("$WeekTag.manual")
        AttemptFile = Join-Path $lock ("$WeekTag.attempts")
    }
}

function Test-ProcessAlive {
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return $false }
    try { $null = Get-Process -Id $ProcessId -ErrorAction Stop; return $true } catch { return $false }
}

# 返回 $null 表示没有正在运行；否则返回锁信息（含是否为失效锁）
function Get-RunningLock {
    param([Parameter(Mandatory)]$Paths)
    if (-not (Test-Path -LiteralPath $Paths.RunningLock)) { return $null }
    $pidVal = 0; $start = ''
    try {
        foreach ($ln in [System.IO.File]::ReadAllLines($Paths.RunningLock)) {
            if ($ln -match '^\s*PID\s*=\s*(\d+)')   { $pidVal = [int]$Matches[1] }
            if ($ln -match '^\s*START\s*=\s*(.+)$') { $start  = $Matches[1].Trim() }
        }
    } catch { }
    $alive = Test-ProcessAlive -ProcessId $pidVal
    return [pscustomobject]@{ ProcessId = $pidVal; Start = $start; Alive = $alive }
}

function Get-AttemptCount {
    param([Parameter(Mandatory)]$Paths)
    if (-not (Test-Path -LiteralPath $Paths.AttemptFile)) { return 0 }
    try {
        $raw = (Get-Content -LiteralPath $Paths.AttemptFile -Raw -ErrorAction Stop).Trim()
        $n = 0
        if ([int]::TryParse($raw, [ref]$n)) { return $n }
    } catch { }
    return 0
}

# ------------------------------------------------------------------ 注册计划任务
function Register-WatcherTask {
    $self = $MyInvocation.MyCommand.Path
    if (-not $self) { $self = Join-Path $BaseDir '脚本\检查并触发周报.ps1' }

    Write-Host ''
    Write-Host '即将注册 Windows 计划任务：' -ForegroundColor Cyan
    Write-Host ("  名称     : {0}" -f $TASK_NAME)
    Write-Host ("  执行日   : 每周五（{0}）" -f $TASK_DAY)
    Write-Host ("  开始时间 : {0}" -f $TASK_AT)
    Write-Host ("  重复间隔 : 每 {0} 分钟" -f $TASK_EVERY_MIN)
    Write-Host ("  持续时间 : {0} 小时（最后一次约 20:00）" -f $TASK_DURATION_H)
    Write-Host ("  命令     : powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $self)
    Write-Host ("  工作目录 : {0}" -f $BaseDir)
    Write-Host ''

    $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
        -Argument ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $self) `
        -WorkingDirectory $BaseDir

    # 每周五 09:00 触发，再挂上「每 15 分钟、持续 11 小时」的重复计划。
    # 每周型触发器本身没有 -RepetitionInterval 参数，用一次性触发器把 Repetition 取过来。
    $weekly = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $TASK_DAY -At $TASK_AT
    $once   = New-ScheduledTaskTrigger -Once -At $TASK_AT `
                -RepetitionInterval (New-TimeSpan -Minutes $TASK_EVERY_MIN) `
                -RepetitionDuration (New-TimeSpan -Hours $TASK_DURATION_H)
    $weekly.Repetition = $once.Repetition

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

    $desc = '地产周报 Agent：每周五 09:00-20:00 每 15 分钟本地检查一次本周公告是否投放完毕；' +
            '满足全部条件才启动 run_weekly.ps1。检查过程不调用任何大模型、不产生 API 请求。'

    try {
        Register-ScheduledTask -TaskName $TASK_NAME -Action $action -Trigger $weekly `
            -Settings $settings -Description $desc -Force | Out-Null
        Write-Host ''
        Write-Host ("已注册计划任务：{0}" -f $TASK_NAME) -ForegroundColor Green
        Write-Host ("查看：Get-ScheduledTask -TaskName {0} | Get-ScheduledTaskInfo" -f $TASK_NAME)
        Write-Host ("删除：Unregister-ScheduledTask -TaskName {0} -Confirm:`$false" -f $TASK_NAME)
        return 0
    } catch {
        Write-Host ("注册失败：{0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-Host '（若提示权限不足，请用「以管理员身份运行」的 PowerShell 重试）' -ForegroundColor Yellow
        return 1
    }
}

# ================================================================== 主流程
$startTime = Get-Date

Set-Location -LiteralPath $BaseDir

# 日志文件（按天追加）；DryRun 也写日志，便于留痕
$logDir = Join-Path $BaseDir $LOG_REL
if (-not (Test-Path -LiteralPath $logDir)) {
    try { New-Item -ItemType Directory -Path $logDir -Force | Out-Null } catch { }
}
$script:LogPath = Join-Path $logDir ('watcher_{0}.log' -f $startTime.ToString('yyyyMMdd'))

Write-Host ''
Write-Host '========== 地产周报 Agent · 自动检查触发 ==========' -ForegroundColor Cyan
Write-Log -Level STEP -Message ("===== 检查开始（{0}{1}）=====" -f $WATCHER_VERSION, $(if ($DryRun) { ' · DryRun' } else { '' }))

if ($RegisterTask) {
    Write-Log -Message '模式：注册计划任务（本次不做检查）'
    $rc = Register-WatcherTask
    exit $rc
}

# 结论收集：每一项写进 $checks，最后统一打印
$checks   = New-Object System.Collections.ArrayList
$blockers = New-Object System.Collections.ArrayList

function Add-Check {
    param([string]$Name, [bool]$Pass, [string]$Detail)
    [void]$checks.Add([pscustomobject]@{ Name = $Name; Pass = $Pass; Detail = $Detail })
    if ($Pass) {
        Write-Log -Level OK   -Message ("  ✔ {0}：{1}" -f $Name, $Detail)
    } else {
        Write-Log -Level SKIP -Message ("  ✘ {0}：{1}" -f $Name, $Detail)
        [void]$blockers.Add($Name)
    }
}

# ---------------------------------------------------------------- 0. 时间窗
Write-Log -Level STEP -Message '[0/9] 运行窗口'
$today   = (Get-Date).Date
$nowTime = Get-Date
$isFriday = ([int]$today.DayOfWeek -eq 5)
$cutoff   = $today.AddHours($CutoffTime.Hour).AddMinutes($CutoffTime.Minute)
# 截止时刻那一次检查仍要执行——它就是当天最后一次机会（20:00 能跑通就跑），
# 20:01 起才算过点。
$inWindow = $nowTime -lt $cutoff.AddMinutes(1)
# 当天最后一个检查时段（含截止时刻与前一次 15 分钟）。用于记录
# 「截至截止时间仍未准备完成」——计划任务的最后一次是 20:00，此时窗口还没关，
# 靠这个标记才能把结论写进日志。
$isLastSlot = $nowTime -ge $cutoff.AddMinutes(-$TASK_EVERY_MIN)

Write-Log -Message ("  今天：{0}（{1}）" -f $today.ToString('yyyy-MM-dd'), $today.DayOfWeek)
Write-Log -Message ("  当前：{0}　截止：{1}" -f $nowTime.ToString('HH:mm:ss'), $cutoff.ToString('HH:mm'))

$dayOk = $isFriday -or $AnyDay
Add-Check -Name '运行日' -Pass $dayOk -Detail $(
    if ($isFriday) { '今天是周五' }
    elseif ($AnyDay) { '今天不是周五，但已指定 -AnyDay（人工补跑）' }
    else { '今天不是周五，常规轮询只在周五进行' })
Add-Check -Name '截止时间' -Pass $inWindow -Detail $(
    if ($inWindow) { "未超过 {0}（截止时刻这一次仍会执行）" -f $cutoff.ToString('HH:mm') }
    else { "已过 {0}，当天不再触发" -f $cutoff.ToString('HH:mm') })

# ---------------------------------------------------------------- 1. 目标周次
Write-Log -Level STEP -Message '[1/9] 推算目标周次'
$target = Get-TargetWeek -Forced $Week
Write-Log -Message ("  窗口末日（周五）：{0}" -f $target.Friday.ToString('yyyy-MM-dd'))
Write-Log -Message ("  目标周次：{0}　输入子目录应为 {1}公司公告" -f $target.Tag, $target.WeekTag)
if ($Week) { Write-Log -Level WARN -Message ("  周次由 -Week 强制指定：{0}" -f $Week) }

# ---------------------------------------------------------------- 2/3. 两侧目录
Write-Log -Level STEP -Message '[2/9] 检查两侧最新周次目录'
$dirA  = Join-Path $BaseDir '02_待处理\A股公告'
$dirHK = Join-Path $BaseDir '02_待处理\港股公告'

$aDirs  = @(Get-WeekDirs -Parent $dirA)
$hkDirs = @(Get-WeekDirs -Parent $dirHK)

$aList  = if ($aDirs.Count  -gt 0) { ($aDirs  | ForEach-Object { $_.Name }) -join '、' } else { '（无）' }
$hkList = if ($hkDirs.Count -gt 0) { ($hkDirs | ForEach-Object { $_.Name }) -join '、' } else { '（无）' }
Write-Log -Message ("  A股  目录下的周次：{0}" -f $aList)
Write-Log -Message ("  港股 目录下的周次：{0}" -f $hkList)

$aLatest  = if ($aDirs.Count  -gt 0) { $aDirs[-1]  } else { $null }
$hkLatest = if ($hkDirs.Count -gt 0) { $hkDirs[-1] } else { $null }

Add-Check -Name 'A股目录' -Pass ($null -ne $aLatest) -Detail $(
    if ($aLatest) { "$($aLatest.Name)（$($aLatest.Path)）" } else { '未找到任何 Wnn公司公告 目录' })
Add-Check -Name '港股目录' -Pass ($null -ne $hkLatest) -Detail $(
    if ($hkLatest) { "$($hkLatest.Name)（$($hkLatest.Path)）" } else { '未找到任何 Wnn公司公告 目录' })

$sameWeek = ($aLatest -and $hkLatest -and $aLatest.Number -eq $hkLatest.Number)
Add-Check -Name '两侧同为最新周次' -Pass $sameWeek -Detail $(
    if ($sameWeek) { "两侧都是 $($aLatest.Tag)公司公告" }
    elseif ($aLatest -and $hkLatest) { "A股最新 $($aLatest.Tag) / 港股最新 $($hkLatest.Tag)，不一致" }
    else { '两侧至少一侧没有周次目录' })

if (-not $sameWeek) {
    Write-Log -Level WARN -Message '两侧最新周次不一致或缺失 → 本次不触发。'
    $commonNumber = $null
} else {
    $commonNumber = $aLatest.Number
}

# 与目标周次是否一致（默认必须一致，避免误跑旧周次）
$targetMatch = ($null -ne $commonNumber -and $commonNumber -eq $target.IsoWeek) -or $IgnoreTargetWeek
Add-Check -Name '最新周次 = 本周目标周次' -Pass $targetMatch -Detail $(
    if ($null -eq $commonNumber) { '无共同周次，跳过' }
    elseif ($commonNumber -eq $target.IsoWeek) { "都是 $($target.Tag)" }
    elseif ($IgnoreTargetWeek) { "最新为 W$('{0:d2}' -f $commonNumber)，与目标 $($target.Tag) 不同，但已指定 -IgnoreTargetWeek" }
    else { "最新为 W$('{0:d2}' -f $commonNumber)，而本周目标是 $($target.Tag)（若需补跑旧周次请用 -Week 指定）" })

$weekTagFull = if ($null -ne $commonNumber) { '{0}-W{1:d2}' -f $target.IsoYear, $commonNumber } else { $target.Tag }
$paths = Get-WeekPaths -WeekTag $weekTagFull

# ---------------------------------------------------------------- 4. 非空 + 5. 稳定
Write-Log -Level STEP -Message '[3/9] 检查两侧目录是否非空、文件是否已稳定'
$bothDirs = @()
if ($aLatest)  { $bothDirs += $aLatest.Path }
if ($hkLatest) { $bothDirs += $hkLatest.Path }

$statA  = if ($aLatest)  { Get-DirFileStat -Path $aLatest.Path }  else { [pscustomobject]@{ Exists=$false; Count=0; Newest=$null } }
$statHK = if ($hkLatest) { Get-DirFileStat -Path $hkLatest.Path } else { [pscustomobject]@{ Exists=$false; Count=0; Newest=$null } }

$nonEmpty = ($statA.Count -gt 0) -and ($statHK.Count -gt 0)
Add-Check -Name '两侧非空' -Pass $nonEmpty -Detail (
    'A股 {0} 个文件 / 港股 {1} 个文件' -f $statA.Count, $statHK.Count)

$newestOverall = $null
foreach ($s in @($statA, $statHK)) {
    if ($null -ne $s.Newest -and ($null -eq $newestOverall -or $s.Newest -gt $newestOverall)) {
        $newestOverall = $s.Newest
    }
}
$stable = $false
if ($null -ne $newestOverall) {
    $stable = (((Get-Date) - $newestOverall).TotalMinutes -ge $StableMinutes)
}
Write-Log -Message ("  A股  最新写入：{0}" -f (Format-Age $statA.Newest))
Write-Log -Message ("  港股 最新写入：{0}" -f (Format-Age $statHK.Newest))
Add-Check -Name ("文件稳定 {0} 分钟" -f $StableMinutes) -Pass $stable -Detail $(
    if ($null -eq $newestOverall) { '没有文件可判断' }
    else { "两侧最新写入为 {0}（要求 ≥ {1} 分钟）" -f (Format-Age $newestOverall), $StableMinutes })

# ---------------------------------------------------------------- 6/7/8/9. 状态
Write-Log -Level STEP -Message '[4/9] 检查该周次是否已完成 / 已叫停 / 正在运行 / 已超尝试次数'
Write-Log -Message ("  周次目录：{0}" -f $paths.ReportDir)

# 6. 已完成
$doneExists = (Test-Path -LiteralPath $paths.DoneFlag) -and (-not $IgnoreDone)
Add-Check -Name '该周次尚未完成' -Pass (-not $doneExists) -Detail $(
    if (Test-Path -LiteralPath $paths.DoneFlag) {
        if ($IgnoreDone) { "done.flag 存在，但已指定 -IgnoreDone（测试）" }
        else { "已存在 done.flag（$($paths.DoneFlag)）→ 同一周只正式运行一次，本次不触发" }
    } else { '未发现 done.flag' })

# 7. 人工叫停
$manualExists = Test-Path -LiteralPath $paths.ManualLock
Add-Check -Name '未被人工叫停' -Pass (-not $manualExists) -Detail $(
    if ($manualExists) { "存在 $($paths.ManualLock) → 已停止自动尝试，等待人工检查" } else { '无 .manual 标记' })

# 8. 正在运行
$running = Get-RunningLock -Paths $paths
$isRunning = ($null -ne $running) -and $running.Alive
Add-Check -Name '没有正在运行' -Pass (-not $isRunning) -Detail $(
    if ($null -eq $running) { '无 .running 锁' }
    elseif ($running.Alive) { "PID $($running.ProcessId) 仍在运行（起始 $($running.Start)）" }
    else { "发现失效锁（PID $($running.ProcessId) 已不存在，起始 $($running.Start)）→ 视为未在运行" })

# 9. 尝试次数
$attempts = Get-AttemptCount -Paths $paths
$underCap = ($attempts -lt $MaxRuns)
Add-Check -Name ("尝试次数未达上限（{0}/{1}）" -f $attempts, $MaxRuns) -Pass $underCap -Detail $(
    if ($underCap) { "已尝试 $attempts 次" }
    else { "已连续失败 $attempts 次，达到上限 → 停止自动尝试" })

# ---------------------------------------------------------------- 判定
$allPass = ($blockers.Count -eq 0)

Write-Host ''
Write-Host '-------------------- 检查结论 --------------------' -ForegroundColor Cyan
foreach ($c in $checks) {
    $mark = if ($c.Pass) { '✔' } else { '✘' }
    $color = if ($c.Pass) { 'Green' } else { 'DarkGray' }
    Write-Host ("  {0} {1}：{2}" -f $mark, $c.Name, $c.Detail) -ForegroundColor $color
}
Write-Host ''

if (-not $allPass) {
    Write-Log -Level WARN -Message ("未触发：{0}" -f ($blockers -join '；'))

    # 规则六：截至截止时间仍未准备完成 → 只在日志记录，不自行抓取替代数据。
    #   只有「卡在输入」才记这条；卡在其它原因（已跑过 / 正在跑 / 已叫停 / 非周五 /
    #   过点）不是输入问题，不能记成「未准备完成」。
    #   且只在当天最后一个检查时段记，避免 09:00 起每 15 分钟刷屏。
    $inputBlockers = @('A股目录', '港股目录', '两侧同为最新周次', '两侧非空', '最新周次 = 本周目标周次', "文件稳定 $StableMinutes 分钟")
    $stuckOnInput = @($blockers | Where-Object { $inputBlockers -contains $_ }).Count -gt 0
    if ($stuckOnInput -and $isLastSlot -and -not $doneExists) {
        Write-Log -Level WARN -Message '【本周输入未准备完成，自动整理未触发】'
        Write-Log -Message '  不会自行抓取替代数据。可稍后人工执行：.\脚本\检查并触发周报.ps1'
    }
    if ((-not $inWindow) -and $dayOk) {
        Write-Log -Level WARN -Message ("已过当日截止时间 {0}，今天不再触发。" -f $cutoff.ToString('HH:mm'))
    }

    # 连续失败达上限 → 落 .manual，之后不再自动尝试
    if (-not $underCap -and -not $manualExists -and -not $DryRun) {
        try {
            $msg = @(
                "周次：$weekTagFull"
                "停止时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                "原因：连续失败 $attempts 次（上限 $MaxRuns），已停止自动尝试"
                "日志：$($script:LogPath)"
                "恢复：人工确认无误后删除本文件，或直接人工执行 .\脚本\run_weekly.ps1 -Week $weekTagFull"
            ) -join "`r`n"
            [void](New-LockDir)
            Set-Content -LiteralPath $paths.ManualLock -Value $msg -Encoding UTF8
            Write-Log -Level ERROR -Message ("已创建人工介入标记：{0}" -f $paths.ManualLock)
        } catch {
            Write-Log -Level ERROR -Message ("创建 .manual 失败：{0}" -f $_.Exception.Message)
        }
    } elseif ((-not $underCap) -and $DryRun -and -not $manualExists) {
        Write-Log -Level WARN -Message ("（DryRun）本应创建人工介入标记：{0}" -f $paths.ManualLock)
    }

    Write-Log -Level OK -Message ("===== 检查结束（未触发，耗时 {0:N1}s）=====" -f ((Get-Date) - $startTime).TotalSeconds)
    Write-Host ("本次未触发 Agent。日志：{0}" -f $script:LogPath) -ForegroundColor Yellow
    Write-Host ''
    exit 0
}

# ---------------------------------------------------------------- 触发
Write-Log -Level OK -Message '全部条件满足 → 准备启动 run_weekly.ps1'
$runPath = Join-Path $BaseDir $RUN_WEEKLY
if (-not (Test-Path -LiteralPath $runPath)) {
    Write-Log -Level ERROR -Message ("找不到运行脚本：{0}" -f $runPath)
    exit 1
}

# 传给 run_weekly 的参数：显式给出周次与窗口末日，保证与本次检查用的周次一致
$endDate  = $target.Friday.ToString('yyyy-MM-dd')
$runArgs  = @('-Week', $weekTagFull, '-EndDate', $endDate)
Write-Log -Message ("启动命令：powershell -NoProfile -ExecutionPolicy Bypass -File `"{0}`" {1}" -f $runPath, ($runArgs -join ' '))
Write-Log -Message ("窗口末日：{0}　周次：{1}" -f $endDate, $weekTagFull)

if ($DryRun) {
    Write-Log -Level WARN -Message '-DryRun 已指定：条件全部满足，但**不启动** run_weekly、不调用大模型、不产生 API 请求。'
    Write-Log -Level WARN -Message '（DryRun）本应写入 .running 锁、本应把尝试次数记为 1'
    Write-Log -Level OK   -Message ("===== 检查结束（DryRun，条件已满足，未启动，耗时 {0:N1}s）=====" -f ((Get-Date) - $startTime).TotalSeconds)
    Write-Host 'DryRun：条件全部满足，本应启动 run_weekly.ps1，但未启动。' -ForegroundColor Yellow
    Write-Host ("日志：{0}" -f $script:LogPath) -ForegroundColor Green
    Write-Host ''
    exit 0
}

# 走到这里才真正要写状态，此时才创建锁目录
try { [void](New-LockDir) } catch {
    Write-Log -Level ERROR -Message ("创建锁目录失败：{0}" -f $_.Exception.Message)
    exit 1
}

# 尝试次数先加 1 再启动：中途崩溃也算消耗一次，避免无限重试
$nextAttempt = $attempts + 1
try {
    Set-Content -LiteralPath $paths.AttemptFile -Value "$nextAttempt" -Encoding ASCII
    Write-Log -Message ("尝试次数记为 {0}/{1}" -f $nextAttempt, $MaxRuns)
} catch {
    Write-Log -Level WARN -Message ("尝试次数写入失败：{0}" -f $_.Exception.Message)
}

# 运行锁（记 PID，便于识别失效锁）
try {
    $lockBody = @("PID=$PID", "START=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", "WEEK=$weekTagFull") -join "`r`n"
    Set-Content -LiteralPath $paths.RunningLock -Value $lockBody -Encoding UTF8
    Write-Log -Message ("已创建运行锁：{0}" -f $paths.RunningLock)
} catch {
    Write-Log -Level WARN -Message ("运行锁写入失败：{0}" -f $_.Exception.Message)
}

$rc = 1
try {
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $runPath @runArgs
    $rc = $LASTEXITCODE
    $ErrorActionPreference = $prevEAP
} catch {
    Write-Log -Level ERROR -Message ("启动 run_weekly.ps1 抛出异常：{0}" -f $_.Exception.Message)
    $rc = 1
} finally {
    try { Remove-Item -LiteralPath $paths.RunningLock -Force -ErrorAction SilentlyContinue } catch { }
}

Write-Log -Message ("run_weekly.ps1 退出码：{0}" -f $rc)

# run_weekly 的退出码 0 只代表 claude 跑完；产出是否齐全必须自己再核一遍，
# 否则会把"跑完但没产出"误判为成功而写下 done.flag。
$expected = @(
    '本周公司公告汇总.md'
    '本周公司公告来源清单.csv'
    '本周重要资讯.md'
    'PPT候选内容.md'
    'run_log.md'
    '周报更新说明.md'
    '本周公司公告周报.docx'
    '本周重要资讯.docx'
)
$missing = @()
foreach ($f in $expected) {
    if (-not (Test-Path -LiteralPath (Join-Path $paths.ReportDir $f))) { $missing += $f }
}

if ($rc -eq 0 -and $missing.Count -eq 0) {
    $body = @(
        "周次：$weekTagFull"
        "完成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "窗口：$($target.Friday.AddDays(-6).ToString('yyyy-MM-dd')) ~ $endDate"
        "输入：A股 $($statA.Count) 个 / 港股 $($statHK.Count) 个 文件"
        "输出目录：$($paths.ReportDir)"
        "本次为第 $nextAttempt 次尝试"
    ) -join "`r`n"
    try {
        Set-Content -LiteralPath $paths.DoneFlag -Value $body -Encoding UTF8
        Write-Log -Level OK -Message ("已写入完成标记：{0}" -f $paths.DoneFlag)
        Write-Log -Level OK -Message '同一周次此后不再自动运行。'
    } catch {
        Write-Log -Level ERROR -Message ("写入 done.flag 失败：{0}" -f $_.Exception.Message)
    }
    Write-Log -Level OK -Message ("===== 检查结束（已触发且成功，耗时 {0:N1}s）=====" -f ((Get-Date) - $startTime).TotalSeconds)
    Write-Host ("已触发并成功完成：{0}" -f $weekTagFull) -ForegroundColor Green
    Write-Host ("输出目录：{0}" -f $paths.ReportDir) -ForegroundColor Green
    Write-Host ''
    exit 0
}

Write-Log -Level ERROR -Message 'run_weekly.ps1 未成功完成 → **不写 done.flag**，保留日志与中间产物。'
if ($rc -ne 0)     { Write-Log -Level ERROR -Message ("  原因：run_weekly.ps1 退出码 {0}" -f $rc) }
if ($missing.Count -gt 0) {
    Write-Log -Level ERROR -Message ("  原因：产出缺失 {0} 个：{1}" -f $missing.Count, ($missing -join '、'))
}
Write-Log -Level WARN -Message ("已尝试 {0}/{1} 次；未达上限时下一次 15 分钟检查会再试。" -f $nextAttempt, $MaxRuns)
if ($nextAttempt -ge $MaxRuns) {
    try {
        $msg = @(
            "周次：$weekTagFull"
            "停止时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            "原因：连续失败 $nextAttempt 次（上限 $MaxRuns），已停止自动尝试"
            "run_weekly 退出码：$rc"
            "缺失产出：$(if ($missing.Count -gt 0) { $missing -join '、' } else { '（无）' })"
            "日志：$($script:LogPath)"
            "恢复：人工确认无误后删除本文件，或直接人工执行 .\脚本\run_weekly.ps1 -Week $weekTagFull"
        ) -join "`r`n"
        [void](New-LockDir)
        Set-Content -LiteralPath $paths.ManualLock -Value $msg -Encoding UTF8
        Write-Log -Level ERROR -Message ("已达上限 → 已创建人工介入标记：{0}" -f $paths.ManualLock)
        Write-Log -Level ERROR -Message '之后不再自动尝试，等待人工检查。'
    } catch {
        Write-Log -Level ERROR -Message ("创建 .manual 失败：{0}" -f $_.Exception.Message)
    }
}

Write-Log -Level OK -Message ("===== 检查结束（已触发但失败，耗时 {0:N1}s）=====" -f ((Get-Date) - $startTime).TotalSeconds)
Write-Host ("已触发 {0}，但未成功完成。日志：{1}" -f $weekTagFull, $script:LogPath) -ForegroundColor Red
Write-Host ''
exit 1
