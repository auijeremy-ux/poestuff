<#
    Start-App.ps1

    The Long Path Shortener app window. Opened by "Long Path Shortener.bat" or
    the desktop shortcut. A zip file or folder dropped onto either of those is
    passed in as -Source.

    Uses Windows Forms, which is built into Windows. Written for Windows
    PowerShell 5.1. Keep this file plain ASCII.

    The window only collects choices and shows results. All the work is done by
    LongPathShortener.psm1, run in the background by LongPathShortener.App.psm1
    so the window stays responsive.
#>
param(
    [string]$Source = ''
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

$engineDir = $PSScriptRoot
$toolRoot = Split-Path -Parent $engineDir
$enginePath = Join-Path $engineDir 'LongPathShortener.psm1'
$appModulePath = Join-Path $engineDir 'LongPathShortener.App.psm1'
$readmePath = Join-Path $toolRoot 'README.md'

# ---------------------------------------------------------------------------
# Error handling. The window has no console, so problems are shown in a
# message box and also written to %LOCALAPPDATA%\LongPathShortener\app-errors.log.
# ---------------------------------------------------------------------------

function Write-AppErrorLog {
    param([string]$Text)
    try {
        if ($env:LOCALAPPDATA) { $base = $env:LOCALAPPDATA } else { $base = [System.IO.Path]::GetTempPath() }
        $dir = Join-Path $base 'LongPathShortener'
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        $line = '{0}  {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Text
        [System.IO.File]::AppendAllText((Join-Path $dir 'app-errors.log'), $line + [Environment]::NewLine)
    } catch { }
}

function Show-Message {
    param(
        [string]$Text,
        [string]$Title = 'Long Path Shortener',
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK
    )
    $owner = Get-Variable -Name form -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($null -ne $owner -and -not $owner.IsDisposed -and $owner.Visible) {
        return [System.Windows.Forms.MessageBox]::Show($owner, $Text, $Title, $Buttons, $Icon)
    }
    return [System.Windows.Forms.MessageBox]::Show($Text, $Title, $Buttons, $Icon)
}

function Invoke-UiAction {
    # Runs a button's work and shows any error instead of failing silently.
    param([scriptblock]$Action)
    try {
        & $Action
    } catch {
        Write-AppErrorLog ($_ | Out-String)
        [void](Show-Message ("Something went wrong:`n`n" + $_.Exception.Message) 'Long Path Shortener' ([System.Windows.Forms.MessageBoxIcon]::Error))
    }
}

try {
    Import-Module $enginePath -Force
    Import-Module $appModulePath -Force
} catch {
    Write-AppErrorLog ($_ | Out-String)
    [void](Show-Message ("The app could not start:`n`n" + $_.Exception.Message) 'Long Path Shortener' ([System.Windows.Forms.MessageBoxIcon]::Error))
    exit 1
}

# ---------------------------------------------------------------------------
# State, settings, fonts and colours
# ---------------------------------------------------------------------------

$settingsPath = Get-AppSettingsPath
$settings = Read-AppSettings $settingsPath $toolRoot

$state = @{
    Job            = $null   # the background run, while one is going
    Busy           = $false
    Mode           = ''      # Check or Apply
    RunKey         = ''      # the choices used for the current run
    CheckedKey     = ''      # the choices used for the last successful check
    CheckedFiles   = 0
    Table          = $null   # results shown in the grid
    ReportPath     = ''
    OutputPath     = ''
    NextText       = ''
    TickErrorShown = $false
}
$colorCache = @{}

$fontBase = New-Object System.Drawing.Font('Segoe UI', 9)
$fontBold = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$fontTitle = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$fontSummary = New-Object System.Drawing.Font('Segoe UI', 10.5, [System.Drawing.FontStyle]::Bold)
$colorHint = [System.Drawing.Color]::FromArgb(90, 90, 90)
$colorAccent = [System.Drawing.Color]::FromArgb(0, 95, 170)
$colorGood = [System.Drawing.Color]::FromArgb(16, 124, 16)
$colorWarn = [System.Drawing.Color]::FromArgb(150, 85, 0)
$colorError = [System.Drawing.Color]::FromArgb(190, 35, 25)

# ---------------------------------------------------------------------------
# Small builders for controls
# ---------------------------------------------------------------------------

function New-UiLabel {
    param([string]$Text, $Font = $null, $Color = $null)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 7, 3, 3)
    if ($null -ne $Font) { $l.Font = $Font }
    if ($null -ne $Color) { $l.ForeColor = $Color }
    return $l
}

function New-UiButton {
    param([string]$Text, [int]$MinWidth = 96)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.AutoSize = $true
    $b.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowOnly
    $b.MinimumSize = New-Object System.Drawing.Size($MinWidth, 30)
    $b.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    $b.UseVisualStyleBackColor = $true
    return $b
}

function New-UiFlow {
    $f = New-Object System.Windows.Forms.FlowLayoutPanel
    $f.AutoSize = $true
    $f.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $f.WrapContents = $false
    $f.Margin = New-Object System.Windows.Forms.Padding(0)
    return $f
}

function New-UiTable {
    param([int]$Columns)
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.ColumnCount = $Columns
    $t.AutoSize = $true
    $t.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $t.Dock = [System.Windows.Forms.DockStyle]::Fill
    $t.Margin = New-Object System.Windows.Forms.Padding(0)
    return $t
}

function Add-UiRow {
    param($Table, [string]$SizeType = 'AutoSize', [single]$Value = 0)
    [void]$Table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]$SizeType, $Value)))
}

function Add-UiColumn {
    param($Table, [string]$SizeType = 'AutoSize', [single]$Value = 0)
    [void]$Table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]$SizeType, $Value)))
}

function New-AppIcon {
    # A simple drawn icon: a long line above a short one, like a path being shortened.
    try {
        $bmp = New-Object System.Drawing.Bitmap(32, 32)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::Transparent)
        $brush = New-Object System.Drawing.SolidBrush($colorAccent)
        $g.FillRectangle($brush, 1, 1, 30, 30)
        $g.FillRectangle([System.Drawing.Brushes]::White, 6, 9, 20, 4)
        $g.FillRectangle([System.Drawing.Brushes]::White, 6, 19, 11, 4)
        $brush.Dispose()
        $g.Dispose()
        return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    } catch {
        return $null
    }
}

try {
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # -----------------------------------------------------------------------
    # The window
    # -----------------------------------------------------------------------

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Long Path Shortener'
    $form.Font = $fontBase
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Size = New-Object System.Drawing.Size(1120, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(880, 600)
    $form.BackColor = [System.Drawing.Color]::White
    $appIcon = New-AppIcon
    if ($null -ne $appIcon) { $form.Icon = $appIcon }

    $tips = New-Object System.Windows.Forms.ToolTip
    $tips.AutoPopDelay = 15000

    $main = New-UiTable 1
    $main.AutoSize = $false
    $main.Padding = New-Object System.Windows.Forms.Padding(16, 10, 16, 10)
    Add-UiColumn $main 'Percent' 100
    foreach ($i in 1..6) { Add-UiRow $main 'AutoSize' }
    Add-UiRow $main 'Percent' 100
    Add-UiRow $main 'AutoSize'

    # Heading -----------------------------------------------------------------
    $lblTitle = New-UiLabel 'Long Path Shortener' $fontTitle $colorAccent
    $lblTitle.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
    $lblSubtitle = New-UiLabel 'Makes client zip files and folders safe to put in SharePoint. Your original is never changed.' $null $colorHint
    $lblSubtitle.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 10)
    $main.Controls.Add($lblTitle, 0, 0)
    $main.Controls.Add($lblSubtitle, 0, 1)

    # Steps 1 and 2 -------------------------------------------------------------
    $inputs = New-UiTable 3
    Add-UiColumn $inputs 'AutoSize'
    Add-UiColumn $inputs 'Percent' 100
    Add-UiColumn $inputs 'AutoSize'
    foreach ($i in 1..4) { Add-UiRow $inputs 'AutoSize' }

    $lblSource = New-UiLabel '1.  Client zip file or folder' $fontBold
    $txtSource = New-Object System.Windows.Forms.TextBox
    $txtSource.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtSource.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    $srcButtons = New-UiFlow
    $btnZip = New-UiButton 'Choose zip file...' 120
    $btnFolder = New-UiButton 'Choose folder...' 120
    $srcButtons.Controls.Add($btnZip)
    $srcButtons.Controls.Add($btnFolder)
    $lblSourceHint = New-UiLabel 'Or drag a zip file or folder onto this window.' $null $colorHint
    $lblSourceHint.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 8)

    $lblDest = New-UiLabel '2.  SharePoint folder' $fontBold
    $cmbDest = New-Object System.Windows.Forms.ComboBox
    $cmbDest.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $cmbDest.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cmbDest.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    $btnDest = New-UiButton 'Browse...' 120
    $lblDestHint = New-UiLabel ('Where the files will finally live. Open the synced SharePoint folder in File Explorer, click the address bar, copy the path and paste it here.' + [Environment]::NewLine + 'For example: C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd') $null $colorHint
    $lblDestHint.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 8)
    $lblDestHint.MaximumSize = New-Object System.Drawing.Size(620, 0)

    $inputs.Controls.Add($lblSource, 0, 0)
    $inputs.Controls.Add($txtSource, 1, 0)
    $inputs.Controls.Add($srcButtons, 2, 0)
    $inputs.Controls.Add($lblSourceHint, 1, 1)
    $inputs.SetColumnSpan($lblSourceHint, 2)
    $inputs.Controls.Add($lblDest, 0, 2)
    $inputs.Controls.Add($cmbDest, 1, 2)
    $inputs.Controls.Add($btnDest, 2, 2)
    $inputs.Controls.Add($lblDestHint, 1, 3)
    $inputs.SetColumnSpan($lblDestHint, 2)
    $main.Controls.Add($inputs, 0, 2)

    # Steps 3 and 4 -------------------------------------------------------------
    $actions = New-UiTable 2
    Add-UiColumn $actions 'Percent' 100
    Add-UiColumn $actions 'AutoSize'
    Add-UiRow $actions 'AutoSize'
    $leftActions = New-UiFlow
    $btnCheck = New-UiButton '3.  Check' 130
    $btnCheck.Font = $fontBold
    $btnApply = New-UiButton '4.  Apply changes' 150
    $btnApply.Font = $fontBold
    $btnStop = New-UiButton 'Stop' 90
    $leftActions.Controls.Add($btnCheck)
    $leftActions.Controls.Add($btnApply)
    $leftActions.Controls.Add($btnStop)
    $rightActions = New-UiFlow
    $btnSettings = New-UiButton 'Settings...' 100
    $btnHelp = New-UiButton 'Help' 80
    $rightActions.Controls.Add($btnSettings)
    $rightActions.Controls.Add($btnHelp)
    $rightActions.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $actions.Controls.Add($leftActions, 0, 0)
    $actions.Controls.Add($rightActions, 1, 0)
    $main.Controls.Add($actions, 0, 3)

    $tips.SetToolTip($btnCheck, 'Works out the new names and lists them below. Nothing is copied.')
    $tips.SetToolTip($btnApply, 'Copies or unzips the files with the new names into the output folder. Your original is not changed.')
    $tips.SetToolTip($btnStop, 'Stops the current job. Files already copied stay in the output folder and are listed in the log.')

    # Summary -------------------------------------------------------------------
    $summary = New-Object System.Windows.Forms.FlowLayoutPanel
    $summary.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $summary.WrapContents = $false
    $summary.AutoSize = $true
    $summary.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 4)
    $lblSummary = New-UiLabel 'Choose the zip file or folder and the SharePoint folder, then click Check.' $fontSummary
    $lblNext = New-UiLabel 'Nothing is copied until you click Apply changes.' $null $colorHint
    $lblNext.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 3)
    $summary.Controls.Add($lblSummary)
    $summary.Controls.Add($lblNext)
    $main.Controls.Add($summary, 0, 4)

    # Filter --------------------------------------------------------------------
    $filterBar = New-UiFlow
    $filterBar.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)
    $chkProblems = New-Object System.Windows.Forms.CheckBox
    $chkProblems.Text = 'Show only problems'
    $chkProblems.AutoSize = $true
    $chkProblems.Margin = New-Object System.Windows.Forms.Padding(3, 6, 18, 3)
    $lblSearch = New-UiLabel 'Search:'
    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Width = 260
    $txtSearch.Margin = New-Object System.Windows.Forms.Padding(3, 4, 12, 3)
    $lblCount = New-UiLabel '' $null $colorHint
    $filterBar.Controls.Add($chkProblems)
    $filterBar.Controls.Add($lblSearch)
    $filterBar.Controls.Add($txtSearch)
    $filterBar.Controls.Add($lblCount)
    $main.Controls.Add($filterBar, 0, 5)

    # Results grid --------------------------------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    $grid.RowHeadersVisible = $false
    $grid.AutoGenerateColumns = $false
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $grid.ColumnHeadersDefaultCellStyle.Font = $fontBold
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(204, 228, 247)
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    try {
        # Smoother scrolling. DoubleBuffered is not public, so it is set by reflection.
        $grid.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'NonPublic, Instance').SetValue($grid, $true, $null)
    } catch { }

    function New-GridColumn {
        param([string]$Name, [string]$Header, [single]$Weight, [int]$MinWidth)
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name = $Name
        $c.DataPropertyName = $Name
        $c.HeaderText = $Header
        $c.FillWeight = $Weight
        $c.MinimumWidth = $MinWidth
        $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        return $c
    }
    $colStatus = New-GridColumn 'Status' 'Status' 15 150
    $colOriginal = New-GridColumn 'Original' 'Original path (inside the SharePoint folder)' 34 160
    $colNew = New-GridColumn 'New' 'New path' 34 160
    $colLength = New-GridColumn 'NewLength' 'Length' 7 64
    $colLength.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
    $colLength.ToolTipText = 'Characters in the full new path, including the SharePoint folder.'
    $colChanges = New-GridColumn 'Changes' 'What changed' 22 120
    [void]$grid.Columns.Add($colStatus)
    [void]$grid.Columns.Add($colOriginal)
    [void]$grid.Columns.Add($colNew)
    [void]$grid.Columns.Add($colLength)
    [void]$grid.Columns.Add($colChanges)
    $main.Controls.Add($grid, 0, 6)

    # Footer ---------------------------------------------------------------------
    $footer = New-UiTable 3
    Add-UiColumn $footer 'AutoSize'
    Add-UiColumn $footer 'Percent' 100
    Add-UiColumn $footer 'AutoSize'
    Add-UiRow $footer 'AutoSize'
    $footer.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Width = 220
    $progress.Height = 18
    $progress.Margin = New-Object System.Windows.Forms.Padding(3, 8, 10, 3)
    $progress.MarqueeAnimationSpeed = 30
    $progress.Visible = $false
    $lblStatus = New-UiLabel 'Ready.' $null $colorHint
    $lblStatus.AutoSize = $false
    $lblStatus.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblStatus.AutoEllipsis = $true
    $footerButtons = New-UiFlow
    $btnReport = New-UiButton 'Open report' 110
    $btnOutput = New-UiButton 'Open output folder' 140
    $footerButtons.Controls.Add($btnReport)
    $footerButtons.Controls.Add($btnOutput)
    $footer.Controls.Add($progress, 0, 0)
    $footer.Controls.Add($lblStatus, 1, 0)
    $footer.Controls.Add($footerButtons, 2, 0)
    $main.Controls.Add($footer, 0, 7)

    $form.Controls.Add($main)

    # -----------------------------------------------------------------------
    # Behaviour
    # -----------------------------------------------------------------------

    function Get-RunKey {
        # Everything that affects the result. Apply is only allowed while this
        # matches the last check, so staff always apply exactly what they saw.
        $parts = @(
            $txtSource.Text.Trim().Trim('"'), $cmbDest.Text.Trim().Trim('"'),
            $settings.OutputFolder, $settings.MaxPathLength, $settings.SafetyMargin,
            $settings.ExpandNestedZips, $settings.IncludeSystemFiles, $settings.AbbreviationsCsv
        )
        return ($parts -join '|')
    }

    function Update-Buttons {
        $busy = [bool]$state.Busy
        $checkedCurrent = ($state.CheckedKey -ne '') -and ($state.CheckedKey -eq (Get-RunKey))
        $btnCheck.Enabled = -not $busy
        $btnApply.Enabled = (-not $busy) -and $checkedCurrent -and ($state.CheckedFiles -gt 0)
        $btnStop.Enabled = $busy
        $btnSettings.Enabled = -not $busy
        foreach ($c in @($txtSource, $cmbDest, $btnZip, $btnFolder, $btnDest)) { $c.Enabled = -not $busy }
        $btnReport.Enabled = (-not $busy) -and ($state.ReportPath -ne '')
        $btnOutput.Enabled = (-not $busy) -and ($state.OutputPath -ne '')
        if (-not $busy -and $state.CheckedKey -ne '' -and -not $checkedCurrent) {
            $lblNext.Text = 'You have changed something since the check. Click Check again before applying.'
        } elseif ($state.NextText -ne '') {
            $lblNext.Text = $state.NextText
        }
    }

    function Update-DestinationList {
        $current = $cmbDest.Text
        $cmbDest.BeginUpdate()
        $cmbDest.Items.Clear()
        foreach ($d in @($settings.RecentDestinations)) { if ($d) { [void]$cmbDest.Items.Add($d) } }
        $cmbDest.EndUpdate()
        $cmbDest.Text = $current
    }

    function Set-Source {
        param([string]$Path)
        $txtSource.Text = $Path.Trim().Trim('"')
        if ($cmbDest.Text.Trim() -eq '') { [void]$cmbDest.Focus() }
    }

    function Update-GridFilter {
        if ($null -eq $state.Table) { $lblCount.Text = ''; return }
        $state.Table.DefaultView.RowFilter = Get-ResultFilter ([bool]$chkProblems.Checked) $txtSearch.Text
        $lblCount.Text = 'Showing {0} of {1}' -f $state.Table.DefaultView.Count, $state.Table.Rows.Count
    }

    function Update-LabelWidths {
        # Long labels wrap to the window width instead of stretching the window.
        $w = [Math]::Max(300, $main.ClientSize.Width - 40)
        $lblSummary.MaximumSize = New-Object System.Drawing.Size($w, 0)
        $lblNext.MaximumSize = New-Object System.Drawing.Size($w, 0)
        $lblSourceHint.MaximumSize = New-Object System.Drawing.Size([Math]::Max(300, $w - 460), 0)
        $lblDestHint.MaximumSize = New-Object System.Drawing.Size([Math]::Max(300, $w - 460), 0)
    }

    function Open-InExplorer {
        param([string]$Path)
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Path)
    }

    function Start-Run {
        param([string]$Mode)
        $src = $txtSource.Text.Trim().Trim('"')
        $dest = $cmbDest.Text.Trim().Trim('"')
        if ($src -eq '') {
            [void](Show-Message 'First choose the client''s zip file or folder (step 1).' 'Long Path Shortener' ([System.Windows.Forms.MessageBoxIcon]::Warning))
            return
        }
        if ($dest -eq '') {
            [void](Show-Message 'Enter the SharePoint folder the files will go into (step 2).' 'Long Path Shortener' ([System.Windows.Forms.MessageBoxIcon]::Warning))
            [void]$cmbDest.Focus()
            return
        }

        # Warn if the output folder would start syncing straight away.
        $syncMatch = ''
        try { $syncMatch = Get-SyncedFolderMatch -OutputFolder $settings.OutputFolder -DestinationPrefix $dest } catch { $syncMatch = '' }
        if ($syncMatch -and $Mode -eq 'Check') {
            $question = ("The output folder`n    {0}`nis inside a OneDrive or SharePoint synced folder:`n    {1}`n`n" +
                "Anything written there starts uploading straight away, including files that still need attention. " +
                "It is safer to use a folder outside OneDrive, such as C:\CL\Out (change it in Settings).`n`nContinue anyway?") -f $settings.OutputFolder, $syncMatch
            $answer = Show-Message $question 'Output folder is synced' ([System.Windows.Forms.MessageBoxIcon]::Warning) ([System.Windows.Forms.MessageBoxButtons]::YesNo)
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        if ($Mode -eq 'Apply') {
            $question = "Copy the files with their new names into:`n    {0}`n`nYour original zip file or folder will not be changed." -f $settings.OutputFolder
            $answer = Show-Message $question 'Apply changes' ([System.Windows.Forms.MessageBoxIcon]::Question) ([System.Windows.Forms.MessageBoxButtons]::OKCancel)
            if ($answer -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }

        $params = @{
            Source             = $src
            DestinationPrefix  = $dest
            OutputFolder       = [string]$settings.OutputFolder
            MaxPathLength      = [int]$settings.MaxPathLength
            SafetyMargin       = [int]$settings.SafetyMargin
            ExpandNestedZips   = [bool]$settings.ExpandNestedZips
            IncludeSystemFiles = [bool]$settings.IncludeSystemFiles
            AllowSyncedOutput  = [bool]($syncMatch -ne '')
        }
        if ($settings.AbbreviationsCsv -and (Test-Path -LiteralPath $settings.AbbreviationsCsv)) { $params.AbbreviationsCsv = [string]$settings.AbbreviationsCsv }
        if ($Mode -eq 'Apply') {
            $params.Apply = $true
            if ($settings.CreateZip) { $params.CreateZip = $true }
        }

        Add-RecentDestination $settings $dest
        try { Save-AppSettings $settings $settingsPath } catch { Write-AppErrorLog ('Could not save settings: ' + $_.Exception.Message) }
        Update-DestinationList

        $state.Mode = $Mode
        $state.RunKey = Get-RunKey
        $state.Busy = $true
        $state.TickErrorShown = $false
        $state.Job = Start-ShortenerJob -Parameters $params -ModulePath $enginePath
        $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progress.Visible = $true
        if ($Mode -eq 'Apply') { $lblStatus.Text = 'Copying...' } else { $lblStatus.Text = 'Checking...' }
        Update-Buttons
        $timer.Start()
    }

    function Complete-WithResult {
        param($Result)
        $table = ConvertTo-ResultTable $Result.Rows $Result.DestinationPrefix
        $state.Table = $table
        $grid.DataSource = $table
        Update-GridFilter

        $sum = Get-ResultSummary $Result
        $lblSummary.Text = $sum.Text
        if ($sum.Level -eq 'Good') { $lblSummary.ForeColor = $colorGood }
        elseif ($sum.Level -eq 'Warning') { $lblSummary.ForeColor = $colorWarn }
        else { $lblSummary.ForeColor = $colorError }
        $state.NextText = $sum.Next
        $lblNext.Text = $sum.Next
        $state.ReportPath = [string]$Result.ReportPath

        if ($state.Mode -eq 'Check') {
            $state.CheckedKey = $state.RunKey
            $state.CheckedFiles = [int]$Result.FilesProcessed
            $state.OutputPath = ''
            $lblStatus.Text = 'Check finished. Nothing has been copied yet.'
            return
        }

        # Applying again would make a second copy, so a new check is needed first.
        $state.CheckedKey = ''
        $state.OutputPath = [string]$Result.FilesFolder
        $lblStatus.Text = 'Finished.'
        $msg = "Finished.`n`nUpload the CONTENTS of this folder to the SharePoint folder:`n    {0}" -f $Result.FilesFolder
        if ($Result.AttentionFolder) {
            $msg += "`n`n{0} file(s) are still too long. They are in:`n    {1}`nSort these out by hand before uploading them (see Help)." -f $Result.FilesFlagged, $Result.AttentionFolder
        }
        if ([int]$Result.Errors -gt 0) {
            $msg += "`n`n{0} file(s) could not be copied. They are marked Error in the list." -f $Result.Errors
        }
        $msg += "`n`nKeep the evidence log with the matter:`n    {0}`n`nOpen the output folder now?" -f $Result.LogPath
        $answer = Show-Message $msg 'Finished' ([System.Windows.Forms.MessageBoxIcon]::Information) ([System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) { Open-InExplorer $Result.FilesFolder }
    }

    function Complete-WithError {
        param([string]$Message)
        if ($Message -like 'Stopped: cancelled*') {
            $lblStatus.Text = 'Stopped.'
            if ($state.Mode -eq 'Apply') {
                $text = "Stopped.`n`nFiles copied before you pressed Stop are in the output folder ({0}) and are listed in the log saved there. Your original was not changed." -f $settings.OutputFolder
            } else {
                $text = 'Stopped. Nothing was copied.'
            }
            [void](Show-Message $text 'Stopped')
        } else {
            $lblStatus.Text = 'Could not finish.'
            Write-AppErrorLog $Message
            [void](Show-Message $Message 'Could not finish' ([System.Windows.Forms.MessageBoxIcon]::Warning))
        }
    }

    function Update-JobProgress {
        # Returns $true while the job is still running.
        $job = $state.Job
        if ($null -eq $job) { return $false }
        $sync = $job.Sync
        if ($sync.Cancel -and -not $sync.Finished) { $lblStatus.Text = 'Stopping...' } else { $lblStatus.Text = [string]$sync.Message }
        $done = [long]$sync.Done
        $total = [long]$sync.Total
        if ($total -gt 0 -and $done -ge 0) {
            if ($progress.Style -ne [System.Windows.Forms.ProgressBarStyle]::Continuous) {
                $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $progress.Maximum = 1000
            }
            $progress.Value = [int][Math]::Min(1000, [Math]::Floor(1000.0 * $done / $total))
        } elseif ($progress.Style -ne [System.Windows.Forms.ProgressBarStyle]::Marquee) {
            $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        }
        if (-not $sync.Finished) { return $true }

        Complete-ShortenerJob $job
        $state.Job = $null
        $state.Busy = $false
        $progress.Visible = $false
        try {
            if ($sync.Error) { Complete-WithError ([string]$sync.Error) } else { Complete-WithResult $sync.Result }
        } finally {
            Update-Buttons
        }
        return $false
    }

    function Show-SettingsDialog {
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = 'Settings'
        $dlg.Font = $fontBase
        $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dlg.MaximizeBox = $false
        $dlg.MinimizeBox = $false
        $dlg.ShowInTaskbar = $false
        $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $dlg.AutoSize = $true
        $dlg.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
        $dlg.Padding = New-Object System.Windows.Forms.Padding(14)
        if ($null -ne $appIcon) { $dlg.Icon = $appIcon }

        $t = New-UiTable 3
        Add-UiColumn $t 'AutoSize'
        Add-UiColumn $t 'AutoSize'
        Add-UiColumn $t 'AutoSize'
        foreach ($i in 1..9) { Add-UiRow $t 'AutoSize' }

        $txtOut = New-Object System.Windows.Forms.TextBox
        $txtOut.Width = 400
        $txtOut.Text = $settings.OutputFolder
        $btnOut = New-UiButton 'Browse...' 90
        $lblOutHint = New-UiLabel 'Renamed files, reports and logs go here. Keep it outside OneDrive and SharePoint.' $null $colorHint
        $lblOutHint.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 8)

        $numMax = New-Object System.Windows.Forms.NumericUpDown
        $numMax.Minimum = 50
        $numMax.Maximum = 400
        $numMax.Width = 80
        $numMax.Value = [decimal]$settings.MaxPathLength
        $flowMax = New-UiFlow
        $flowMax.Controls.Add($numMax)
        $flowMax.Controls.Add((New-UiLabel 'characters. Excel stops at 218, Word at 259.' $null $colorHint))

        $numMargin = New-Object System.Windows.Forms.NumericUpDown
        $numMargin.Minimum = 0
        $numMargin.Maximum = 100
        $numMargin.Width = 80
        $numMargin.Value = [decimal]$settings.SafetyMargin
        $flowMargin = New-UiFlow
        $flowMargin.Controls.Add($numMargin)
        $flowMargin.Controls.Add((New-UiLabel 'characters kept spare below the limit.' $null $colorHint))

        $txtAbbr = New-Object System.Windows.Forms.TextBox
        $txtAbbr.Width = 400
        $txtAbbr.Text = $settings.AbbreviationsCsv
        $btnAbbr = New-UiButton 'Browse...' 90

        $chkNested = New-Object System.Windows.Forms.CheckBox
        $chkNested.Text = 'Unpack zip files found inside the zip or folder (up to 3 levels deep)'
        $chkNested.AutoSize = $true
        $chkNested.Checked = [bool]$settings.ExpandNestedZips
        $chkZip = New-Object System.Windows.Forms.CheckBox
        $chkZip.Text = 'Also make a new zip of the renamed files'
        $chkZip.AutoSize = $true
        $chkZip.Checked = [bool]$settings.CreateZip
        $chkSystem = New-Object System.Windows.Forms.CheckBox
        $chkSystem.Text = 'Keep Mac and Windows housekeeping files (__MACOSX, .DS_Store, Thumbs.db)'
        $chkSystem.AutoSize = $true
        $chkSystem.Checked = [bool]$settings.IncludeSystemFiles

        $buttons = New-UiFlow
        $btnDefaults = New-UiButton 'Restore defaults' 130
        $btnSave = New-UiButton 'Save' 90
        $btnCancel = New-UiButton 'Cancel' 90
        $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $buttons.Controls.Add($btnDefaults)
        $buttons.Controls.Add($btnSave)
        $buttons.Controls.Add($btnCancel)
        $buttons.Anchor = [System.Windows.Forms.AnchorStyles]::Right
        $buttons.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)

        $t.Controls.Add((New-UiLabel 'Output folder' $fontBold), 0, 0)
        $t.Controls.Add($txtOut, 1, 0)
        $t.Controls.Add($btnOut, 2, 0)
        $t.Controls.Add($lblOutHint, 1, 1)
        $t.SetColumnSpan($lblOutHint, 2)
        $t.Controls.Add((New-UiLabel 'Longest path allowed' $fontBold), 0, 2)
        $t.Controls.Add($flowMax, 1, 2)
        $t.SetColumnSpan($flowMax, 2)
        $t.Controls.Add((New-UiLabel 'Safety margin' $fontBold), 0, 3)
        $t.Controls.Add($flowMargin, 1, 3)
        $t.SetColumnSpan($flowMargin, 2)
        $t.Controls.Add((New-UiLabel 'Abbreviations file' $fontBold), 0, 4)
        $t.Controls.Add($txtAbbr, 1, 4)
        $t.Controls.Add($btnAbbr, 2, 4)
        $t.Controls.Add($chkNested, 0, 5)
        $t.SetColumnSpan($chkNested, 3)
        $t.Controls.Add($chkZip, 0, 6)
        $t.SetColumnSpan($chkZip, 3)
        $t.Controls.Add($chkSystem, 0, 7)
        $t.SetColumnSpan($chkSystem, 3)
        $t.Controls.Add($buttons, 0, 8)
        $t.SetColumnSpan($buttons, 3)
        $dlg.Controls.Add($t)
        $dlg.AcceptButton = $btnSave
        $dlg.CancelButton = $btnCancel

        $btnOut.Add_Click({
                Invoke-UiAction {
                    $d = New-Object System.Windows.Forms.FolderBrowserDialog
                    $d.Description = 'Choose where renamed files, reports and logs should go'
                    $d.ShowNewFolderButton = $true
                    if (Test-Path -LiteralPath $txtOut.Text) { $d.SelectedPath = $txtOut.Text }
                    if ($d.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) { $txtOut.Text = $d.SelectedPath }
                    $d.Dispose()
                }
            })
        $btnAbbr.Add_Click({
                Invoke-UiAction {
                    $d = New-Object System.Windows.Forms.OpenFileDialog
                    $d.Title = 'Choose the abbreviations file'
                    $d.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
                    if ($d.ShowDialog($dlg) -eq [System.Windows.Forms.DialogResult]::OK) { $txtAbbr.Text = $d.FileName }
                    $d.Dispose()
                }
            })
        $btnDefaults.Add_Click({
                Invoke-UiAction {
                    $def = Get-DefaultAppSettings $toolRoot
                    $txtOut.Text = $def.OutputFolder
                    $numMax.Value = [decimal]$def.MaxPathLength
                    $numMargin.Value = [decimal]$def.SafetyMargin
                    $txtAbbr.Text = $def.AbbreviationsCsv
                    $chkNested.Checked = [bool]$def.ExpandNestedZips
                    $chkZip.Checked = [bool]$def.CreateZip
                    $chkSystem.Checked = [bool]$def.IncludeSystemFiles
                }
            })
        $dlg.Add_FormClosing({
                param($ctl, $e)
                if ($dlg.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return }
                $problem = ''
                if ($txtOut.Text.Trim() -eq '') { $problem = 'Choose an output folder.' }
                elseif ([int]$numMargin.Value -ge ([int]$numMax.Value - 40)) { $problem = 'The safety margin is too big for that path limit.' }
                elseif ($txtAbbr.Text.Trim() -ne '' -and -not (Test-Path -LiteralPath $txtAbbr.Text.Trim())) { $problem = 'The abbreviations file cannot be found. Leave the box empty to use none.' }
                if ($problem) {
                    [void](Show-Message $problem 'Settings' ([System.Windows.Forms.MessageBoxIcon]::Warning))
                    $e.Cancel = $true
                }
            })

        $result = $dlg.ShowDialog($form)
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $settings.OutputFolder = $txtOut.Text.Trim().Trim('"')
            $settings.MaxPathLength = [int]$numMax.Value
            $settings.SafetyMargin = [int]$numMargin.Value
            $settings.AbbreviationsCsv = $txtAbbr.Text.Trim().Trim('"')
            $settings.ExpandNestedZips = [bool]$chkNested.Checked
            $settings.CreateZip = [bool]$chkZip.Checked
            $settings.IncludeSystemFiles = [bool]$chkSystem.Checked
            Save-AppSettings $settings $settingsPath
        }
        $dlg.Dispose()
    }

    # -----------------------------------------------------------------------
    # Events
    # -----------------------------------------------------------------------

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 150
    $timer.Add_Tick({
            $timer.Stop()
            $running = $false
            try {
                $running = Update-JobProgress
            } catch {
                # Keep polling so the window is not left stuck, but only report once.
                $running = ($null -ne $state.Job)
                Write-AppErrorLog ($_ | Out-String)
                if (-not $state.TickErrorShown) {
                    $state.TickErrorShown = $true
                    [void](Show-Message ('Something went wrong while showing progress:' + "`n`n" + $_.Exception.Message) 'Long Path Shortener' ([System.Windows.Forms.MessageBoxIcon]::Error))
                }
            }
            if ($running) { $timer.Start() }
        })

    $btnZip.Add_Click({
            Invoke-UiAction {
                $d = New-Object System.Windows.Forms.OpenFileDialog
                $d.Title = 'Choose the client''s zip file'
                $d.Filter = 'Zip files (*.zip)|*.zip|All files (*.*)|*.*'
                $d.CheckFileExists = $true
                if ($d.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { Set-Source $d.FileName }
                $d.Dispose()
            }
        })
    $btnFolder.Add_Click({
            Invoke-UiAction {
                $d = New-Object System.Windows.Forms.FolderBrowserDialog
                $d.Description = 'Choose the client''s folder'
                $d.ShowNewFolderButton = $false
                if ($d.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { Set-Source $d.SelectedPath }
                $d.Dispose()
            }
        })
    $btnDest.Add_Click({
            Invoke-UiAction {
                $d = New-Object System.Windows.Forms.FolderBrowserDialog
                $d.Description = 'Choose the synced SharePoint folder these files will go into'
                $d.ShowNewFolderButton = $false
                $current = $cmbDest.Text.Trim().Trim('"')
                if ($current -and (Test-Path -LiteralPath $current)) { $d.SelectedPath = $current }
                if ($d.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $cmbDest.Text = $d.SelectedPath }
                $d.Dispose()
            }
        })

    $btnCheck.Add_Click({ Invoke-UiAction { Start-Run 'Check' } })
    $btnApply.Add_Click({ Invoke-UiAction { Start-Run 'Apply' } })
    $btnStop.Add_Click({
            Invoke-UiAction {
                if ($null -ne $state.Job) {
                    $state.Job.Sync.Cancel = $true
                    $lblStatus.Text = 'Stopping...'
                    $btnStop.Enabled = $false
                }
            }
        })
    $btnSettings.Add_Click({ Invoke-UiAction { Show-SettingsDialog; Update-Buttons } })
    $btnHelp.Add_Click({
            Invoke-UiAction {
                if (Test-Path -LiteralPath $readmePath) {
                    Start-Process -FilePath 'notepad.exe' -ArgumentList ('"{0}"' -f $readmePath)
                } else {
                    [void](Show-Message ('The guide was not found at ' + $readmePath))
                }
            }
        })
    $btnReport.Add_Click({ Invoke-UiAction { if ($state.ReportPath) { Start-Process -FilePath $state.ReportPath } } })
    $btnOutput.Add_Click({ Invoke-UiAction { if ($state.OutputPath) { Open-InExplorer $state.OutputPath } } })

    $txtSource.Add_TextChanged({ Invoke-UiAction { Update-Buttons } })
    $cmbDest.Add_TextChanged({ Invoke-UiAction { Update-Buttons } })
    $chkProblems.Add_CheckedChanged({ Invoke-UiAction { Update-GridFilter } })
    $txtSearch.Add_TextChanged({ Invoke-UiAction { Update-GridFilter } })

    $grid.Add_CellFormatting({
            param($ctl, $e)
            try {
                if ($e.RowIndex -lt 0) { return }
                $view = $grid.Rows[$e.RowIndex].DataBoundItem
                if ($null -eq $view) { return }
                $status = [string]$view.Row['Status']
                if (-not $colorCache.ContainsKey($status)) {
                    $rgb = Get-StatusRgb $status
                    $colorCache[$status] = [System.Drawing.Color]::FromArgb($rgb[0], $rgb[1], $rgb[2])
                }
                $e.CellStyle.BackColor = $colorCache[$status]
                if ($e.ColumnIndex -eq 0) { $e.CellStyle.Font = $fontBold }
            } catch { }
        })
    $grid.Add_CellDoubleClick({
            param($ctl, $e)
            Invoke-UiAction {
                if ($e.RowIndex -lt 0) { return }
                $row = $grid.Rows[$e.RowIndex].DataBoundItem.Row
                $newFull = [string]$row['NewFull']
                if ($newFull -eq '') { $newFull = '(not copied)' }
                $text = "Status: {0}`n`nOriginal path ({1} characters):`n{2}`n`nNew path ({3} characters):`n{4}`n`nWhat changed:`n{5}" -f `
                    $row['Status'], $row['OriginalLength'], $row['OriginalFull'], $row['NewLength'], $newFull, $row['Changes']
                [void](Show-Message $text 'File details')
            }
        })

    # Drag and drop works anywhere on the window.
    $onDragEnter = {
        param($ctl, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop) -and -not $state.Busy) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        } else {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::None
        }
    }
    $onDragDrop = {
        param($ctl, $e)
        Invoke-UiAction {
            $paths = @($e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop))
            if ($paths.Count -gt 0 -and -not $state.Busy) { Set-Source ([string]$paths[0]) }
        }
    }
    function Enable-DropTarget {
        param($Control)
        $Control.AllowDrop = $true
        $Control.Add_DragEnter($onDragEnter)
        $Control.Add_DragDrop($onDragDrop)
        foreach ($child in $Control.Controls) { Enable-DropTarget $child }
    }
    Enable-DropTarget $form

    $form.Add_Shown({
            Invoke-UiAction {
                Update-LabelWidths
                Update-DestinationList
                if ($Source) { Set-Source $Source } else { [void]$txtSource.Focus() }
                Update-Buttons
            }
        })
    $form.Add_Resize({ try { Update-LabelWidths } catch { } })
    $form.Add_FormClosing({
            param($ctl, $e)
            if ($state.Busy) {
                [void](Show-Message 'Please wait for the current job to finish, or click Stop first.' 'Still working')
                $e.Cancel = $true
            }
        })

    Update-Buttons
    [System.Windows.Forms.Application]::Run($form)
} catch {
    Write-AppErrorLog ($_ | Out-String)
    [void](Show-Message ("The app stopped because of a problem:`n`n" + $_.Exception.Message) 'Long Path Shortener' ([System.Windows.Forms.MessageBoxIcon]::Error))
    exit 1
}
