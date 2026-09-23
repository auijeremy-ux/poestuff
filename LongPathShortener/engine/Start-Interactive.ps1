<#
    Start-Interactive.ps1

    The step-by-step prompts behind "Shorten Long Paths.bat":
      1. ask for the zip or folder (unless it was dragged onto the .bat)
      2. ask for the destination SharePoint folder
      3. run a dry run and open the report
      4. ask "Apply these changes? (Y/N)"
      5. copy or extract with the new names and open the output folder

    Exit codes: 0 finished or cancelled, 2 stopped with a message.
#>
param(
    [string]$Source = ''
)

# ---------------------------------------------------------------------------
# Settings. IT or a confident user can change these.
# ---------------------------------------------------------------------------
$OutputFolder = 'C:\CL\Out'
$MaxPathLength = 218
$SafetyMargin = 10
$ExpandNestedZips = $true
$CreateZip = $false
$AbbreviationsCsv = Join-Path (Split-Path -Parent $PSScriptRoot) 'Abbreviations.csv'
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

function Read-CleanAnswer {
    param([string]$Prompt)
    $a = Read-Host $Prompt
    if ($null -eq $a) { return '' }
    return $a.Trim().Trim('"').Trim()
}

function Wait-ForEnter {
    Write-Host ''
    [void](Read-Host 'Press Enter to close this window')
}

function Select-FolderDialog {
    # Built-in Windows folder picker. Returns '' if cancelled or unavailable.
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = 'Choose the synced SharePoint folder these files will go into'
        $d.ShowNewFolderButton = $false
        if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $d.SelectedPath }
    } catch { }
    return ''
}

try {
    Import-Module (Join-Path $PSScriptRoot 'LongPathShortener.psm1') -Force

    Write-Host ''
    Write-Host '  LONG PATH SHORTENER' -ForegroundColor Cyan
    Write-Host '  Makes client folders and zip files safe to put in SharePoint.'
    Write-Host '  Your original zip file or folder is never changed.'
    Write-Host ''

    if (-not $Source) {
        $Source = Read-CleanAnswer '  Drag the zip file or folder into this window, then press Enter'
    }
    if (-not $Source) { throw 'No zip file or folder was given, so nothing was done.' }
    Write-Host ('  Source: ' + $Source)
    Write-Host ''

    # Remember the last destination for this user (stored on this computer only).
    $stateFile = ''
    $last = ''
    if ($env:LOCALAPPDATA) {
        $stateFile = Join-Path (Join-Path $env:LOCALAPPDATA 'LongPathShortener') 'last-destination.txt'
        try { if (Test-Path -LiteralPath $stateFile) { $last = ([string](Get-Content -LiteralPath $stateFile -Raw)).Trim() } } catch { $last = '' }
    }

    Write-Host '  Where will these files finally live?'
    Write-Host '  Open the synced SharePoint folder in File Explorer, click the address bar,'
    Write-Host '  copy the path and paste it here. For example:'
    Write-Host '    C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd'
    Write-Host '  Or type B and press Enter to pick the folder instead.'
    if ($last) { Write-Host ('  Press Enter on its own to use the last one: ' + $last) }
    $prefix = Read-CleanAnswer '  Destination folder'
    if ($prefix -eq 'B' -or $prefix -eq 'b') { $prefix = Select-FolderDialog }
    if (-not $prefix) { $prefix = $last }
    if (-not $prefix) { throw 'No destination folder was given, so nothing was done.' }
    Write-Host ('  Destination: ' + $prefix)
    if ($stateFile) {
        try {
            [void](New-Item -ItemType Directory -Force -Path (Split-Path -Parent $stateFile))
            Set-Content -LiteralPath $stateFile -Value $prefix -Encoding UTF8
        } catch { }
    }

    $common = @{
        Source            = $Source
        DestinationPrefix = $prefix
        OutputFolder      = $OutputFolder
        MaxPathLength     = $MaxPathLength
        SafetyMargin      = $SafetyMargin
        ExpandNestedZips  = [bool]$ExpandNestedZips
    }
    if ($AbbreviationsCsv -and (Test-Path -LiteralPath $AbbreviationsCsv)) { $common.AbbreviationsCsv = $AbbreviationsCsv }

    Write-Host ''
    Write-Host '  STEP 1 OF 2: checking. Nothing is copied or renamed yet.' -ForegroundColor Cyan
    $dry = Invoke-PathShortener @common

    if ($dry.FilesProcessed -eq 0) {
        Write-Host '  No files were found to process. See the report for details.' -ForegroundColor Yellow
    }

    Write-Host '  Opening the report in Excel...'
    try { Invoke-Item -LiteralPath $dry.ReportPath } catch { Write-Host ('  Could not open it automatically. It is here: ' + $dry.ReportPath) }
    Write-Host ''
    Write-Host '  In the report, NewPath shows where each file will end up and Status shows'
    Write-Host '  what will happen to it. Close Excel when you have finished looking.'
    if ($dry.FilesFlagged -gt 0) {
        Write-Host ('  {0} file(s) cannot be shortened enough. They will be copied to a separate' -f $dry.FilesFlagged) -ForegroundColor Yellow
        Write-Host '  "Needs attention" folder for you to sort out by hand (see the README).' -ForegroundColor Yellow
    }
    Write-Host ''

    do {
        $answer = (Read-CleanAnswer '  Apply these changes? (Y/N)').ToUpperInvariant()
    } while ($answer -ne 'Y' -and $answer -ne 'N')

    if ($answer -eq 'N') {
        Write-Host '  OK. Nothing was copied or renamed.'
        Wait-ForEnter
        exit 0
    }

    $applyParams = $common.Clone()
    $applyParams.Apply = $true
    if ($CreateZip) { $applyParams.CreateZip = $true }
    # The synced-folder question was already answered in step 1.
    if ($dry.OutputIsSynced) { $applyParams.AllowSyncedOutput = $true }

    Write-Host ''
    Write-Host '  STEP 2 OF 2: copying files with their new names...' -ForegroundColor Cyan
    $res = Invoke-PathShortener @applyParams

    Write-Host '  Finished.' -ForegroundColor Green
    Write-Host '  Upload the CONTENTS of this folder to the SharePoint folder:'
    Write-Host ('    ' + $res.FilesFolder)
    if ($res.AttentionFolder) {
        Write-Host '  These files still need attention before they can go in SharePoint:' -ForegroundColor Yellow
        Write-Host ('    ' + $res.AttentionFolder) -ForegroundColor Yellow
    }
    Write-Host '  Keep the evidence log with the matter:'
    Write-Host ('    ' + $res.LogPath)
    try { Invoke-Item -LiteralPath $res.FilesFolder } catch { }
    Wait-ForEnter
    exit 0
} catch {
    Write-Host ''
    Write-Host ('  ' + $_.Exception.Message) -ForegroundColor Red
    Wait-ForEnter
    exit 2
}
