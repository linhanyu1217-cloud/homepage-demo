Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$AgentRoot = "C:\Users\15258\Desktop\地产周报Agent"
$LogDir = Join-Path $AgentRoot "日志"
$OutputDir = Join-Path $AgentRoot "05_输出"

$form = New-Object System.Windows.Forms.Form
$form.Text = "地产周报 Agent 状态"
$form.Size = New-Object System.Drawing.Size(820, 560)
$form.StartPosition = "CenterScreen"

$font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
$form.Font = $font

$title = New-Object System.Windows.Forms.Label
$title.Text = "地产周报 Agent 运行状态"
$title.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 16, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(20, 15)
$title.Size = New-Object System.Drawing.Size(500, 35)
$form.Controls.Add($title)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 60)
$statusLabel.Size = New-Object System.Drawing.Size(760, 28)
$statusLabel.Text = "正在检测..."
$form.Controls.Add($statusLabel)

$weekLabel = New-Object System.Windows.Forms.Label
$weekLabel.Location = New-Object System.Drawing.Point(20, 92)
$weekLabel.Size = New-Object System.Drawing.Size(760, 25)
$weekLabel.Text = "当前周次：-"
$form.Controls.Add($weekLabel)

$stepLabel = New-Object System.Windows.Forms.Label
$stepLabel.Location = New-Object System.Drawing.Point(20, 122)
$stepLabel.Size = New-Object System.Drawing.Size(760, 25)
$stepLabel.Text = "当前进度：-"
$form.Controls.Add($stepLabel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 152)
$progress.Size = New-Object System.Drawing.Size(760, 28)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$form.Controls.Add($progress)

$percentLabel = New-Object System.Windows.Forms.Label
$percentLabel.Location = New-Object System.Drawing.Point(20, 184)
$percentLabel.Size = New-Object System.Drawing.Size(760, 25)
$percentLabel.Text = "0%"
$form.Controls.Add($percentLabel)

$processLabel = New-Object System.Windows.Forms.Label
$processLabel.Location = New-Object System.Drawing.Point(20, 216)
$processLabel.Size = New-Object System.Drawing.Size(760, 60)
$processLabel.Text = "运行进程：-"
$form.Controls.Add($processLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(20, 285)
$logBox.Size = New-Object System.Drawing.Size(760, 190)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "立即刷新"
$refreshButton.Location = New-Object System.Drawing.Point(20, 485)
$refreshButton.Size = New-Object System.Drawing.Size(110, 32)
$form.Controls.Add($refreshButton)

$openOutputButton = New-Object System.Windows.Forms.Button
$openOutputButton.Text = "打开输出目录"
$openOutputButton.Location = New-Object System.Drawing.Point(145, 485)
$openOutputButton.Size = New-Object System.Drawing.Size(130, 32)
$form.Controls.Add($openOutputButton)

function Get-LatestLog {
    if (-not (Test-Path $LogDir)) { return $null }
    Get-ChildItem $LogDir -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Get-LatestWeek {
    if (-not (Test-Path $OutputDir)) { return "-" }
    $dir = Get-ChildItem $OutputDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{4}-W\d{2}$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($dir) { return $dir.Name }
    return "-"
}

function Get-AgentProcesses {
    $items = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match '^(claude|node|python|pythonw|powershell|pwsh)(\.exe)?$' -and
            (
                $_.CommandLine -match '地产周报Agent' -or
                $_.CommandLine -match 'run_weekly' -or
                $_.CommandLine -match 'claude\s+-p' -or
                $_.CommandLine -match '检查并触发周报'
            )
        }

    if (-not $items) { return @() }

    foreach ($p in $items) {
        $proc = Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            Name = $p.Name
            Id = $p.ProcessId
            CPU = if ($proc) { [math]::Round($proc.CPU, 1) } else { 0 }
            CommandLine = $p.CommandLine
        }
    }
}

function Get-ProgressFromLog([string[]]$lines) {
    $step = 0
    $total = 6
    $label = "等待开始"

    foreach ($line in $lines) {
        if ($line -match '\[STEP\]\s*\[(\d+)\/(\d+)\]\s*(.*)$') {
            $step = [int]$matches[1]
            $total = [int]$matches[2]
            $label = $matches[3].Trim()
        }
    }

    $percent = 0
    if ($total -gt 0) {
        $percent = [math]::Min(100, [math]::Max(0, [math]::Round(($step / $total) * 100)))
    }

    if ($lines -match '运行完成|任务完成|done|success|已完成') {
        if ($step -ge $total) { $percent = 100 }
    }

    [PSCustomObject]@{
        Step = $step
        Total = $total
        Label = $label
        Percent = $percent
    }
}

function Refresh-Status {
    $week = Get-LatestWeek
    $weekLabel.Text = "当前周次：$week"

    $latestLog = Get-LatestLog
    $lines = @()
    if ($latestLog) {
        try {
            $lines = @(Get-Content $latestLog.FullName -Tail 120 -ErrorAction Stop)
        } catch {
            $lines = @("读取日志失败：$($_.Exception.Message)")
        }
    }

    $prog = Get-ProgressFromLog $lines
    $progress.Value = [int]$prog.Percent
    $percentLabel.Text = "$($prog.Percent)%"
    $stepLabel.Text = "当前进度：[$($prog.Step)/$($prog.Total)] $($prog.Label)"

    $procs = @(Get-AgentProcesses)
    if ($procs.Count -gt 0) {
        $statusLabel.Text = "状态：正在运行"
        $procText = ($procs | ForEach-Object {
            "$($_.Name)  PID=$($_.Id)  CPU=$($_.CPU)s"
        }) -join "    "
        $processLabel.Text = "运行进程：$procText"
    } else {
        if ($latestLog -and $latestLog.LastWriteTime -gt (Get-Date).AddMinutes(-3)) {
            $statusLabel.Text = "状态：最近有活动，当前未检测到Agent主进程"
        } else {
            $statusLabel.Text = "状态：空闲 / 等待新一周输入"
        }
        $processLabel.Text = "运行进程：未检测到 claude / python / run_weekly 活动进程"
    }

    if ($latestLog) {
        $header = "最新日志：$($latestLog.Name)    更新时间：$($latestLog.LastWriteTime)`r`n"
        $logBox.Text = $header + (($lines | Select-Object -Last 35) -join "`r`n")
        $logBox.SelectionStart = $logBox.Text.Length
        $logBox.ScrollToCaret()
    } else {
        $logBox.Text = "尚未找到日志文件。"
    }
}

$refreshButton.Add_Click({ Refresh-Status })
$openOutputButton.Add_Click({
    if (Test-Path $OutputDir) {
        Start-Process explorer.exe $OutputDir
    }
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 2000
$timer.Add_Tick({ Refresh-Status })
$timer.Start()

$form.Add_Shown({ Refresh-Status })
[void]$form.ShowDialog()
