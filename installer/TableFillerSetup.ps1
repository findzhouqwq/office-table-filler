#Requires -Version 5.1
<#
    表格批量填写工具 - 图形化安装程序
    ------------------------------------------------------------
    * 标准向导式界面（欢迎 → 安装位置 → 附加任务 → 安装 → 完成）
    * 支持自定义安装目录
    * 可选创建桌面 / 开始菜单 / 任务栏快捷方式
    * 程序本体与依赖均自动下载：镜像一 → 镜像二 → 官方
    * 无控制台窗口（配合 Install.vbs 或 ps2exe -noConsole 使用）
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$script:AppName = '表格批量填写工具'
$script:AppId = 'TableFiller'
$script:Version = '1.1.0'
$script:DefaultDir = Join-Path $env:LOCALAPPDATA 'Programs\TableFiller'
$script:LogPath = Join-Path $env:TEMP ('TableFiller-setup-' + [Guid]::NewGuid().ToString('N') + '.log')
$script:LogPos = 0
$script:Job = $null
$script:JobError = ''
$script:InstallSucceeded = $false
$script:InstallDir = $script:DefaultDir

$script:Accent = [System.Drawing.Color]::FromArgb(0, 120, 215)
$script:AccentDark = [System.Drawing.Color]::FromArgb(0, 90, 170)
$script:SidebarColor = [System.Drawing.Color]::FromArgb(31, 45, 61)
$script:TextMuted = [System.Drawing.Color]::FromArgb(100, 100, 100)
$script:Success = [System.Drawing.Color]::FromArgb(24, 140, 80)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function New-Label {
    param(
        [string]$Text = '',
        [int]$X = 0, [int]$Y = 0, [int]$W = 100, [int]$H = 20,
        [int]$Size = 9,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($W, $H)
    $label.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', $Size, $Style)
    $label.ForeColor = $Color
    return $label
}

function New-Button {
    param(
        [string]$Text = '按钮',
        [int]$X = 0, [int]$Y = 0, [int]$W = 90, [int]$H = 32,
        [bool]$Primary = $false
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($W, $H)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    if ($Primary) {
        $button.BackColor = $script:Accent
        $button.ForeColor = [System.Drawing.Color]::White
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.FlatAppearance.BorderSize = 0
    }
    return $button
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "$($script:AppName) 安装程序"
$form.ClientSize = New-Object System.Drawing.Size(700, 480)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$iconPath = Join-Path $scriptRoot 'TableFiller.ico'
if (Test-Path -LiteralPath $iconPath) {
    try { $form.Icon = New-Object System.Drawing.Icon($iconPath) } catch { }
}

# --- sidebar ---------------------------------------------------------------
$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Dock = [System.Windows.Forms.DockStyle]::Left
$sidebar.Width = 190
$sidebar.BackColor = $script:SidebarColor
$form.Controls.Add($sidebar)

if (Test-Path -LiteralPath $iconPath) {
    try {
        $logo = New-Object System.Windows.Forms.PictureBox
        $logo.Image = $form.Icon.ToBitmap()
        $logo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $logo.Size = New-Object System.Drawing.Size(72, 72)
        $logo.Location = New-Object System.Drawing.Point(20, 30)
        $sidebar.Controls.Add($logo)
    } catch { }
}

$sideTitle = New-Label -Text "表格批量`r`n填写工具" -X 20 -Y 118 -W 150 -H 70 -Size 15 -Style Bold -Color White
$sidebar.Controls.Add($sideTitle)

$sideSub = New-Label -Text "Office 数据批量填表`r`nWindows 桌面版" -X 20 -Y 196 -W 150 -H 46 -Size 9 -Color ([System.Drawing.Color]::FromArgb(190, 205, 220))
$sidebar.Controls.Add($sideSub)

$sideVer = New-Label -Text "版本 $($script:Version)" -X 20 -Y 430 -W 150 -H 20 -Size 8 -Color ([System.Drawing.Color]::FromArgb(150, 170, 190))
$sidebar.Controls.Add($sideVer)

# --- header ----------------------------------------------------------------
$header = New-Label -Text '欢迎使用' -X 220 -Y 24 -W 450 -H 32 -Size 15 -Style Bold
$form.Controls.Add($header)

$subHeader = New-Label -Text '' -X 220 -Y 58 -W 450 -H 38 -Size 9 -Color $script:TextMuted
$form.Controls.Add($subHeader)

$accentLine = New-Object System.Windows.Forms.Panel
$accentLine.Location = New-Object System.Drawing.Point(220, 96)
$accentLine.Size = New-Object System.Drawing.Size(452, 2)
$accentLine.BackColor = [System.Drawing.Color]::FromArgb(225, 230, 236)
$form.Controls.Add($accentLine)

# --- content ---------------------------------------------------------------
$content = New-Object System.Windows.Forms.Panel
$content.Location = New-Object System.Drawing.Point(220, 108)
$content.Size = New-Object System.Drawing.Size(452, 292)
$content.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($content)

$pages = @()

# page 0: welcome
$pageWelcome = New-Object System.Windows.Forms.Panel
$pageWelcome.Dock = [System.Windows.Forms.DockStyle]::Fill
$welcome = New-Label -X 0 -Y 0 -W 440 -H 260 -Size 9
$welcome.Text = @"
欢迎使用“表格批量填写工具”。

本向导将帮助您完成安装，主要步骤包括：

   • 自动下载程序本体与运行依赖
   • 下载地址优先使用国内镜像，失败时自动切换备用镜像和官方源
   • 可自定义安装位置
   • 可选择创建桌面、开始菜单和任务栏快捷方式

本程序用于图片 OCR 填表、文字填表以及 Excel / Word / PPT 跨文件批量填表。

点击“下一步”继续。
"@
$pageWelcome.Controls.Add($welcome)
$pages += $pageWelcome

# page 1: install directory
$pageDir = New-Object System.Windows.Forms.Panel
$pageDir.Dock = [System.Windows.Forms.DockStyle]::Fill

$lblDir = New-Label -Text '安装到：' -X 0 -Y 4 -W 440 -H 22 -Size 9
$pageDir.Controls.Add($lblDir)

$txtDir = New-Object System.Windows.Forms.TextBox
$txtDir.Location = New-Object System.Drawing.Point(0, 30)
$txtDir.Size = New-Object System.Drawing.Size(344, 28)
$txtDir.Text = $script:DefaultDir
$pageDir.Controls.Add($txtDir)

$btnBrowse = New-Button -Text '浏览…' -X 356 -Y 29 -W 88 -H 30
$pageDir.Controls.Add($btnBrowse)

$dirNote = New-Label -X 0 -Y 76 -W 440 -H 120 -Size 9 -Color $script:TextMuted
$dirNote.Text = @"
建议安装到当前用户目录，这样无需管理员权限。

安装过程需要创建独立运行环境并下载约 500 MB 依赖，
请确保网络可用且有足够磁盘空间。

安装程序不会删除目标目录中已有的其他文件。
"@
$pageDir.Controls.Add($dirNote)
$pages += $pageDir

# page 2: tasks
$pageTasks = New-Object System.Windows.Forms.Panel
$pageTasks.Dock = [System.Windows.Forms.DockStyle]::Fill

$lblTasks = New-Label -Text '快捷方式选项' -X 0 -Y 0 -W 440 -H 24 -Size 10 -Style Bold
$pageTasks.Controls.Add($lblTasks)

$chkDesktop = New-Object System.Windows.Forms.CheckBox
$chkDesktop.Text = '创建桌面快捷方式'
$chkDesktop.Location = New-Object System.Drawing.Point(4, 38)
$chkDesktop.Size = New-Object System.Drawing.Size(400, 26)
$chkDesktop.Checked = $true
$pageTasks.Controls.Add($chkDesktop)

$chkStartMenu = New-Object System.Windows.Forms.CheckBox
$chkStartMenu.Text = '创建开始菜单快捷方式'
$chkStartMenu.Location = New-Object System.Drawing.Point(4, 74)
$chkStartMenu.Size = New-Object System.Drawing.Size(400, 26)
$chkStartMenu.Checked = $true
$pageTasks.Controls.Add($chkStartMenu)

$chkTaskbar = New-Object System.Windows.Forms.CheckBox
$chkTaskbar.Text = '固定到任务栏（Windows 快捷栏）'
$chkTaskbar.Location = New-Object System.Drawing.Point(4, 110)
$chkTaskbar.Size = New-Object System.Drawing.Size(400, 26)
$chkTaskbar.Checked = $false
$pageTasks.Controls.Add($chkTaskbar)

$taskNote = New-Label -X 24 -Y 144 -W 410 -H 80 -Size 8 -Color $script:TextMuted
$taskNote.Text = '说明：Windows 10/11 对自动固定任务栏有限制，部分系统可能无法自动完成；' +
                 '若未生效，可在开始菜单中右键程序选择“固定到任务栏”。'
$pageTasks.Controls.Add($taskNote)
$pages += $pageTasks

# page 3: progress
$pageProgress = New-Object System.Windows.Forms.Panel
$pageProgress.Dock = [System.Windows.Forms.DockStyle]::Fill

$lblStatus = New-Label -Text '准备安装…' -X 0 -Y 0 -W 440 -H 22 -Size 9
$pageProgress.Controls.Add($lblStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(0, 28)
$progressBar.Size = New-Object System.Drawing.Size(444, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$pageProgress.Controls.Add($progressBar)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(0, 58)
$logBox.Size = New-Object System.Drawing.Size(444, 226)
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)
$logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logBox.Font = New-Object System.Drawing.Font('Consolas', 8)
$logBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$pageProgress.Controls.Add($logBox)
$pages += $pageProgress

# page 4: finish
$pageFinish = New-Object System.Windows.Forms.Panel
$pageFinish.Dock = [System.Windows.Forms.DockStyle]::Fill

$finishTitle = New-Label -Text '安装完成' -X 0 -Y 6 -W 440 -H 34 -Size 16 -Style Bold -Color $script:Success
$pageFinish.Controls.Add($finishTitle)

$finishText = New-Label -X 0 -Y 52 -W 440 -H 110 -Size 9
$finishText.Text = @"
表格批量填写工具已成功安装到您的电脑。

您现在可以从桌面、开始菜单或任务栏启动程序。
如需卸载，请打开“设置 → 应用 → 已安装的应用”，
找到“表格批量填写工具”后选择卸载。
"@
$pageFinish.Controls.Add($finishText)

$finishPath = New-Label -X 0 -Y 166 -W 440 -H 40 -Size 8 -Color $script:TextMuted
$pageFinish.Controls.Add($finishPath)

$chkLaunch = New-Object System.Windows.Forms.CheckBox
$chkLaunch.Text = '立即启动 表格批量填写工具'
$chkLaunch.Location = New-Object System.Drawing.Point(4, 214)
$chkLaunch.Size = New-Object System.Drawing.Size(400, 26)
$chkLaunch.Checked = $true
$pageFinish.Controls.Add($chkLaunch)
$pages += $pageFinish

foreach ($page in $pages) {
    $page.Visible = $false
    $content.Controls.Add($page)
}

# --- footer ----------------------------------------------------------------
$footerLine = New-Object System.Windows.Forms.Panel
$footerLine.Location = New-Object System.Drawing.Point(220, 412)
$footerLine.Size = New-Object System.Drawing.Size(452, 1)
$footerLine.BackColor = [System.Drawing.Color]::FromArgb(225, 230, 236)
$form.Controls.Add($footerLine)

$btnCancel = New-Button -Text '取消' -X 590 -Y 426 -W 82 -H 32
$form.Controls.Add($btnCancel)

$btnNext = New-Button -Text '下一步 >' -X 500 -Y 426 -W 82 -H 32 -Primary $true
$form.Controls.Add($btnNext)

$btnBack = New-Button -Text '< 上一步' -X 410 -Y 426 -W 82 -H 32
$form.Controls.Add($btnBack)

# ---------------------------------------------------------------------------
# Page switching
# ---------------------------------------------------------------------------
$script:PageIndex = 0

function Show-Page {
    param([int]$Index)
    $script:PageIndex = $Index
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $pages[$i].Visible = ($i -eq $Index)
    }
    switch ($Index) {
        0 {
            $header.Text = "欢迎使用 $($script:AppName) 安装程序"
            $subHeader.Text = '此向导将引导您完成安装。'
        }
        1 {
            $header.Text = '选择安装位置'
            $subHeader.Text = '请选择程序安装目录，建议使用当前用户目录以避免管理员权限。'
        }
        2 {
            $header.Text = '选择附加任务'
            $subHeader.Text = '请选择安装期间要创建的快捷方式。'
        }
        3 {
            $header.Text = '正在安装'
            $subHeader.Text = '请稍候，安装程序正在下载并配置所需组件。'
        }
        4 {
            $header.Text = '安装完成'
            $subHeader.Text = '表格批量填写工具已成功安装到您的电脑。'
        }
    }
    $btnBack.Enabled = ($Index -eq 1 -or $Index -eq 2)
    $btnBack.Visible = ($Index -ne 4)
    $btnCancel.Visible = ($Index -ne 4)
    $btnNext.Visible = ($Index -ne 3)
    if ($Index -eq 2) {
        $btnNext.Text = '安装'
    } elseif ($Index -eq 4) {
        $btnNext.Text = '完成'
    } else {
        $btnNext.Text = '下一步 >'
    }
    $btnNext.Enabled = -not ($script:Job -and $script:Job.State -eq 'Running')
}

function Add-Log {
    param([string]$Text)
    if (-not $logBox.IsDisposed) {
        $logBox.AppendText($Text + "`r`n")
        $logBox.SelectionStart = $logBox.TextLength
        $logBox.ScrollToCaret()
    }
}

function Handle-LogLine {
    param([string]$Line)
    switch -Regex ($Line) {
        '^PROGRESS\|(\d+)\|(.*)$' {
            $progressBar.Value = [Math]::Max(0, [Math]::Min(100, [int]$Matches[1]))
            $lblStatus.Text = $Matches[2]
            Add-Log $Matches[2]
        }
        '^STATUS\|(.*)$' { Add-Log $Matches[1] }
        '^ERROR\|(.*)$' {
            $script:JobError = $Matches[1]
            Add-Log ('错误：' + $Matches[1])
        }
        '^DONE\|(.*)$' { $script:InstallSucceeded = $true }
        '^PIP\|(.*)$' { Add-Log '    ' + $Matches[1] }
        default { Add-Log $Line }
    }
}

function Read-NewLog {
    if (-not (Test-Path -LiteralPath $script:LogPath)) { return }
    try {
        $stream = [System.IO.File]::Open($script:LogPath, 'Open', 'Read', 'ReadWrite')
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $content = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()
    } catch {
        return
    }
    if ($content.Length -le $script:LogPos) { return }
    $chunk = $content.Substring($script:LogPos)
    if (-not $chunk.EndsWith("`n")) {
        $lastNewline = $chunk.LastIndexOf("`n")
        if ($lastNewline -lt 0) { return }
        $chunk = $chunk.Substring(0, $lastNewline + 1)
    }
    $script:LogPos += $chunk.Length
    foreach ($line in ($chunk -split "`r?`n")) {
        if ($line.Trim().Length -gt 0) {
            Handle-LogLine $line.TrimEnd("`r")
        }
    }
}

# ---------------------------------------------------------------------------
# Installation job
# ---------------------------------------------------------------------------
$jobScript = {
    param($InstallDir, $Desktop, $StartMenu, $Taskbar, $LogPath, $Version)

    $ErrorActionPreference = 'Stop'
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    function Write-Log([string]$Line) {
        [System.IO.File]::AppendAllText($LogPath, $Line + "`r`n", $utf8)
    }
    function Report([int]$Percent, [string]$Message) {
        Write-Log ("PROGRESS|{0}|{1}" -f $Percent, $Message)
    }
    function Log([string]$Message) {
        Write-Log ("STATUS|" + $Message)
    }
    $script:JobFailed = $false
    function Fail([string]$Message) {
        $script:JobFailed = $true
        Write-Log ("ERROR|" + $Message)
        throw $Message
    }

    $appSources = @(
        @{ Name = '镜像一 (ghfast)'; Url = 'https://ghfast.top/https://github.com/findzhouqwq/office-table-filler/archive/refs/heads/main.zip' },
        @{ Name = '镜像二 (gh-proxy)'; Url = 'https://gh-proxy.com/https://github.com/findzhouqwq/office-table-filler/archive/refs/heads/main.zip' },
        @{ Name = 'GitHub 官方'; Url = 'https://github.com/findzhouqwq/office-table-filler/archive/refs/heads/main.zip' }
    )
    $pythonSources = @(
        @{ Name = '华为云镜像'; Url = 'https://mirrors.huaweicloud.com/python/3.12.8/python-3.12.8-amd64.exe' },
        @{ Name = '阿里云镜像'; Url = 'https://mirrors.aliyun.com/python-release/windows/python-3.12.8-amd64.exe' },
        @{ Name = 'Python 官方'; Url = 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe' }
    )
    $pipIndexes = @(
        @{ Name = '清华 PyPI 镜像'; Url = 'https://pypi.tuna.tsinghua.edu.cn/simple' },
        @{ Name = '阿里云 PyPI 镜像'; Url = 'https://mirrors.aliyun.com/pypi/simple/' },
        @{ Name = 'PyPI 官方'; Url = 'https://pypi.org/simple' }
    )

    function Download-WithFallback {
        param($Sources, [string]$Destination, [string]$Label)
        foreach ($source in $Sources) {
            try {
                Log ("正在从 {0} 下载{1}…" -f $source['Name'], $Label)
                $ProgressPreference = 'SilentlyContinue'
                Invoke-WebRequest -Uri $source['Url'] -OutFile $Destination -UseBasicParsing -TimeoutSec 180
                if ((Test-Path -LiteralPath $Destination) -and (Get-Item -LiteralPath $Destination).Length -gt 1024) {
                    Log ("已从 {0} 下载完成。" -f $source['Name'])
                    return
                }
                throw '文件大小异常'
            } catch {
                Log ("{0} 下载失败：{1}" -f $source['Name'], $_.Exception.Message)
                Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
            }
        }
        Fail ("{0}：所有下载源均失败，请检查网络后重试。" -f $Label)
    }

    function Find-Python {
        $candidates = New-Object System.Collections.Generic.List[string]
        foreach ($v in @('312', '311', '310', '313')) {
            $candidates.Add((Join-Path $env:LOCALAPPDATA "Programs\Python\Python$v\python.exe"))
            $candidates.Add((Join-Path $env:ProgramFiles "Python$v\python.exe"))
        }
        $pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
        if ($pyLauncher) {
            foreach ($v in @('3.12', '3.11', '3.10', '3.13')) {
                try {
                    $versionArg = '-' + $v
                    $path = & $pyLauncher.Source $versionArg -c "import sys; print(sys.executable)" 2>$null
                    if ($LASTEXITCODE -eq 0 -and $path) {
                        $candidates.Insert(0, $path.Trim())
                    }
                } catch { }
            }
        }
        $pythonCmd = Get-Command python.exe -ErrorAction SilentlyContinue
        if ($pythonCmd) { $candidates.Add($pythonCmd.Source) }
        foreach ($candidate in $candidates) {
            if (-not (Test-Path -LiteralPath $candidate)) { continue }
            if ($candidate -like '*\WindowsApps\*') { continue }
            try {
                & $candidate -c "import sys; sys.exit(0 if (3,10) <= sys.version_info[:2] <= (3,13) else 1)" *> $null
                if ($LASTEXITCODE -eq 0) { return $candidate }
            } catch { }
        }
        return $null
    }

    $tempDir = Join-Path $env:TEMP ('TableFiller-install-' + [Guid]::NewGuid().ToString('N'))
    try {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        # 1. install directory
        Report 3 '正在准备安装目录…'
        if (-not (Test-Path -LiteralPath $InstallDir)) {
            New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
        }

        # 2. download application archive
        Report 8 '正在下载程序本体…'
        $zipPath = Join-Path $tempDir 'app.zip'
        Download-WithFallback $appSources $zipPath '程序本体'

        # 3. extract application
        Report 28 '正在解压程序本体…'
        $extractDir = Join-Path $tempDir 'app'
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force
        $mainFile = Get-ChildItem -LiteralPath $extractDir -Filter 'main.py' -Recurse | Select-Object -First 1
        if (-not $mainFile) {
            Fail '安装包内容不完整：未找到 main.py。'
        }
        $sourceRoot = $mainFile.Directory.FullName

        Report 38 '正在复制程序文件…'
        try {
            Copy-Item -Path (Join-Path $sourceRoot '*') -Destination $InstallDir -Recurse -Force
        } catch {
            Fail ('复制程序文件失败，请先关闭正在运行的表格批量填写工具。' + $_.Exception.Message)
        }
        $iconSource = Join-Path $InstallDir 'installer\TableFiller.ico'
        $iconTarget = Join-Path $InstallDir 'TableFiller.ico'
        if (Test-Path -LiteralPath $iconSource) {
            Copy-Item -LiteralPath $iconSource -Destination $iconTarget -Force
        }
        Remove-Item -LiteralPath (Join-Path $InstallDir 'installer') -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath (Join-Path $InstallDir 'tests') -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath (Join-Path $InstallDir '__pycache__') -Recurse -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path -LiteralPath (Join-Path $InstallDir 'app'))) {
            Fail '复制后的程序目录不完整。'
        }

        # 4. Python
        Report 46 '正在检查 Python 运行环境…'
        $python = Find-Python
        if (-not $python) {
            Report 52 '未找到 Python 3.10-3.13，正在下载 Python 3.12…'
            $pythonInstaller = Join-Path $tempDir 'python-3.12.8-amd64.exe'
            Download-WithFallback $pythonSources $pythonInstaller 'Python 3.12'
            Report 58 '正在安装 Python 3.12（当前用户，无需管理员）…'
            $process = Start-Process -FilePath $pythonInstaller `
                -ArgumentList '/quiet', 'InstallAllUsers=0', 'PrependPath=1', 'Include_test=0', 'Include_launcher=1', 'SimpleInstall=1' `
                -Wait -PassThru -NoNewWindow
            if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
                Fail ("Python 安装失败，退出码 {0}。" -f $process.ExitCode)
            }
            $python = Find-Python
            if (-not $python) {
                Fail 'Python 安装完成后仍未找到可用的 python.exe。'
            }
        }
        Log ("使用 Python：{0}" -f $python)

        # 5. virtual environment
        Report 64 '正在创建独立运行环境…'
        $venvDir = Join-Path $InstallDir '.venv'
        & $python -m venv $venvDir
        if ($LASTEXITCODE -ne 0) {
            Fail '创建虚拟环境失败，请确认安装目录可写。'
        }
        $venvPython = Join-Path $venvDir 'Scripts\python.exe'
        if (-not (Test-Path -LiteralPath $venvPython)) {
            Fail '虚拟环境创建不完整，未找到 python.exe。'
        }

        # 6. dependencies with index fallback
        Report 70 '正在安装运行依赖，这一步可能需要几分钟…'
        $requirements = Join-Path $InstallDir 'requirements.txt'
        if (-not (Test-Path -LiteralPath $requirements)) {
            Fail '安装目录中缺少 requirements.txt。'
        }
        $pipLog = Join-Path $tempDir 'pip.log'
        $installed = $false
        foreach ($index in $pipIndexes) {
            Report 72 ("正在从 {0} 安装依赖…" -f $index['Name'])
            $output = & $venvPython -m pip install --disable-pip-version-check --no-input `
                --timeout 90 --retries 2 -i $index['Url'] -r $requirements 2>&1
            $output | Out-File -LiteralPath $pipLog -Encoding UTF8 -Append
            if ($LASTEXITCODE -eq 0) {
                $installed = $true
                Log ("依赖已从 {0} 安装完成。" -f $index['Name'])
                break
            }
            Log ("{0} 安装依赖失败，自动切换到下一个源…" -f $index['Name'])
        }
        if (-not $installed) {
            $tail = Get-Content -LiteralPath $pipLog -Tail 30 -ErrorAction SilentlyContinue
            foreach ($line in $tail) { Write-Log ('PIP|' + $line) }
            Fail '依赖安装失败，请检查网络后重试。'
        }

        # 7. shortcuts, uninstaller, registry
        Report 92 '正在创建快捷方式与卸载信息…'
        $pythonw = Join-Path $venvDir 'Scripts\pythonw.exe'
        $mainPy = Join-Path $InstallDir 'main.py'
        $icon = Join-Path $InstallDir 'TableFiller.ico'
        $shortcutName = '表格批量填写工具.lnk'
        $launchShortcut = Join-Path $InstallDir $shortcutName

        function New-AppShortcut {
            param([string]$Path)
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($Path)
            $shortcut.TargetPath = $pythonw
            $shortcut.Arguments = '"' + $mainPy + '"'
            $shortcut.WorkingDirectory = $InstallDir
            if (Test-Path -LiteralPath $icon) {
                $shortcut.IconLocation = $icon + ',0'
            }
            $shortcut.Description = '表格批量填写工具'
            $shortcut.Save()
        }

        New-AppShortcut $launchShortcut

        if ($Desktop) {
            try {
                $desktopLnk = Join-Path ([Environment]::GetFolderPath('Desktop')) $shortcutName
                New-AppShortcut $desktopLnk
                Log '已创建桌面快捷方式。'
            } catch {
                Log ('创建桌面快捷方式失败：' + $_.Exception.Message)
            }
        }
        if ($StartMenu) {
            try {
                $menuDir = Join-Path ([Environment]::GetFolderPath('Programs')) '表格批量填写工具'
                New-Item -ItemType Directory -Path $menuDir -Force | Out-Null
                New-AppShortcut (Join-Path $menuDir $shortcutName)
                Log '已创建开始菜单快捷方式。'
            } catch {
                Log ('创建开始菜单快捷方式失败：' + $_.Exception.Message)
            }
        }
        if ($Taskbar) {
            try {
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace($InstallDir)
                $item = $folder.ParseName($shortcutName)
                $verb = $item.Verbs() | Where-Object {
                    $_.Name -match 'taskbar|任务栏' -and $_.Name -notmatch 'unpin|取消'
                } | Select-Object -First 1
                if ($verb) {
                    $verb.DoIt()
                    Log '已请求固定到任务栏；若未出现，请手动固定。'
                } else {
                    Log '当前系统不支持自动固定任务栏，请在开始菜单中手动固定。'
                }
            } catch {
                Log ('固定到任务栏失败：' + $_.Exception.Message)
            }
        }

        # uninstaller
        $uninstallContent = @'
param()
Add-Type -AssemblyName System.Windows.Forms
$dir = '__INSTALL_DIR__'
$answer = [System.Windows.Forms.MessageBox]::Show(
    "确定要卸载 表格批量填写工具 吗？`r`n`r`n程序目录：`r`n$dir",
    '卸载 表格批量填写工具',
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question)
if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { exit }

$shortcutName = '表格批量填写工具.lnk'
$pinnedShortcut = Join-Path $env:APPDATA ('Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\' + $shortcutName)
Remove-Item -LiteralPath $pinnedShortcut -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('Desktop')) $shortcutName) -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('Programs')) '表格批量填写工具') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\TableFiller' -Recurse -Force -ErrorAction SilentlyContinue

$command = '/c timeout /t 1 >nul & rmdir /s /q "' + $dir + '"'
Start-Process -FilePath $env:ComSpec -ArgumentList $command -WindowStyle Hidden
'@
        $uninstallContent = $uninstallContent.Replace('__INSTALL_DIR__', $InstallDir.Replace("'", "''"))
        $uninstallPath = Join-Path $InstallDir 'uninstall.ps1'
        [System.IO.File]::WriteAllText($uninstallPath, $uninstallContent, (New-Object System.Text.UTF8Encoding($true)))

        $uninstallCmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $uninstallPath + '"'
        $registryKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\TableFiller'
        New-Item -Path $registryKey -Force | Out-Null
        Set-ItemProperty -Path $registryKey -Name DisplayName -Value '表格批量填写工具'
        Set-ItemProperty -Path $registryKey -Name DisplayVersion -Value $Version
        Set-ItemProperty -Path $registryKey -Name Publisher -Value 'findzhouqwq'
        Set-ItemProperty -Path $registryKey -Name InstallLocation -Value $InstallDir
        Set-ItemProperty -Path $registryKey -Name DisplayIcon -Value $icon
        Set-ItemProperty -Path $registryKey -Name UninstallString -Value $uninstallCmd
        Set-ItemProperty -Path $registryKey -Name NoModify -Value 1 -Type DWord
        Set-ItemProperty -Path $registryKey -Name NoRepair -Value 1 -Type DWord
        Log '已注册卸载信息。'

        Report 100 '安装完成。'
        Write-Log ('DONE|' + $InstallDir)
    } catch {
        if (-not $script:JobFailed) {
            Write-Log ('ERROR|' + $_.Exception.Message)
        }
    } finally {
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 400
$timer.Add_Tick({
    Read-NewLog
    if ($script:Job -and $script:Job.State -in @('Completed', 'Failed', 'Stopped')) {
        $timer.Stop()
        $state = $script:Job.State
        try { Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue } catch { }
        $script:Job = $null
        if ($script:InstallSucceeded) {
            $finishPath.Text = '安装位置：' + $script:InstallDir
            Show-Page 4
        } else {
            $message = $script:JobError
            if (-not $message) { $message = "安装未完成（任务状态：$state）。" }
            [System.Windows.Forms.MessageBox]::Show(
                $message + "`r`n`r`n请查看安装日志后重试。",
                '安装失败',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            Show-Page 2
        }
    }
})

$btnBack.Add_Click({
    if ($script:PageIndex -gt 0) { Show-Page ($script:PageIndex - 1) }
})

$btnCancel.Add_Click({
    if ($script:Job -and $script:Job.State -eq 'Running') {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            '安装尚未完成，确定要取消吗？',
            '取消安装',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        try { Stop-Job -Job $script:Job -ErrorAction SilentlyContinue } catch { }
        try { Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue } catch { }
        $script:Job = $null
    }
    $form.Close()
})

$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = '请选择安装目录'
    $dialog.ShowNewFolderButton = $true
    if (Test-Path -LiteralPath $txtDir.Text) {
        $dialog.SelectedPath = $txtDir.Text
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDir.Text = $dialog.SelectedPath
    }
})

$btnNext.Add_Click({
    if ($script:PageIndex -eq 0) {
        Show-Page 1
        return
    }
    if ($script:PageIndex -eq 1) {
        $dir = $txtDir.Text.Trim()
        if (-not $dir) {
            [System.Windows.Forms.MessageBox]::Show('请选择安装目录。', '提示', 'OK', 'Warning') | Out-Null
            return
        }
        if ($dir -match '^[A-Za-z]:\\?$') {
            [System.Windows.Forms.MessageBox]::Show('请不要直接安装到磁盘根目录。', '提示', 'OK', 'Warning') | Out-Null
            return
        }
        try { $dir = [System.IO.Path]::GetFullPath($dir) } catch { }
        if (Test-Path -LiteralPath $dir) {
            $items = Get-ChildItem -LiteralPath $dir -Force -ErrorAction SilentlyContinue
            if ($items.Count -gt 0) {
                $answer = [System.Windows.Forms.MessageBox]::Show(
                    ('目录 "' + $dir + '" 已存在且不为空。' + "`r`n`r`n" + '继续安装会覆盖同名文件，是否继续？'),
                    '确认目录',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }
        }
        $txtDir.Text = $dir
        Show-Page 2
        return
    }
    if ($script:PageIndex -eq 2) {
        $script:InstallDir = $txtDir.Text.Trim()
        $script:JobError = ''
        $script:InstallSucceeded = $false
        $script:LogPos = 0
        Remove-Item -LiteralPath $script:LogPath -Force -ErrorAction SilentlyContinue
        $logBox.Clear()
        $progressBar.Value = 0
        $lblStatus.Text = '准备安装…'
        $script:Job = Start-Job -ScriptBlock $jobScript -ArgumentList `
            $script:InstallDir, $chkDesktop.Checked, $chkStartMenu.Checked, $chkTaskbar.Checked, `
            $script:LogPath, $script:Version
        Show-Page 3
        $timer.Start()
        return
    }
    if ($script:PageIndex -eq 4) {
        if ($chkLaunch.Checked) {
            try {
                $pythonw = Join-Path $script:InstallDir '.venv\Scripts\pythonw.exe'
                $mainPy = Join-Path $script:InstallDir 'main.py'
                Start-Process -FilePath $pythonw -ArgumentList ('"' + $mainPy + '"') -WorkingDirectory $script:InstallDir
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    '启动程序失败：' + $_.Exception.Message, '提示', 'OK', 'Warning') | Out-Null
            }
        }
        $form.Close()
    }
})

$form.Add_FormClosing({
    param($sender, $eventArgs)
    if ($script:Job -and $script:Job.State -eq 'Running') {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            '安装仍在进行，确定要退出吗？',
            '退出安装',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $eventArgs.Cancel = $true
            return
        }
        try { Stop-Job -Job $script:Job -ErrorAction SilentlyContinue } catch { }
        try { Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue } catch { }
    }
    $timer.Stop()
})

Show-Page 0
[void]$form.ShowDialog()
