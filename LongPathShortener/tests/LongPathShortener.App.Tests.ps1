<#
    LongPathShortener.App.Tests.ps1

    Pester 3.4 tests for the app window's support code: settings, the results
    table and filter, the summary text, running the engine in the background,
    stopping a run, and checks on the window script itself.

    The window can only be drawn on Windows, so Start-App.ps1 is checked
    statically here: it must parse, and every command and variable it uses must
    exist. Uses only synthetic data.
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$toolRoot = Split-Path -Parent $here
$engineDir = Join-Path $toolRoot 'engine'
$enginePath = Join-Path $engineDir 'LongPathShortener.psm1'
Import-Module $enginePath -Force
Import-Module (Join-Path $engineDir 'LongPathShortener.App.psm1') -Force

$isWin = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
$sep = [string][System.IO.Path]::DirectorySeparatorChar
$workRoot = [System.IO.Path]::GetTempPath() + 'lpsapp-' + [guid]::NewGuid().ToString('N').Substring(0, 6)
$fx = & (Join-Path $here 'New-TestFixtures.ps1') -Path ($workRoot + $sep + 'fx')
$prefix = 'C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd'

function Get-ErrorMessage {
    param([scriptblock]$Script)
    try { & $Script | Out-Null; return '' } catch { return $_.Exception.Message }
}

function Wait-ForJob {
    param($Job, [int]$Seconds = 180)
    $limit = (Get-Date).AddSeconds($Seconds)
    while (-not $Job.Sync.Finished -and (Get-Date) -lt $limit) { Start-Sleep -Milliseconds 100 }
    Complete-ShortenerJob $Job
}

function New-TestRow {
    param([string]$Rel, [string]$NewRel, [string]$Status, [string]$Rules = '')
    $orig = $prefix + '\' + $Rel
    if ($NewRel) { $new = $prefix + '\' + $NewRel; $newLen = $new.Length } else { $new = ''; $newLen = '' }
    return [pscustomobject]@{
        OriginalPath = $orig; NewPath = $new; OriginalLength = $orig.Length; NewLength = $newLen
        RulesApplied = $Rules; Status = $Status
    }
}

# ---------------------------------------------------------------------------

Describe 'App settings' {
    $settingsFile = $workRoot + $sep + 'settings' + $sep + 'settings.json'

    It 'uses the defaults when there is no settings file' {
        $s = Read-AppSettings $settingsFile $toolRoot
        $s.OutputFolder | Should Be 'C:\CL\Out'
        $s.MaxPathLength | Should Be 218
        $s.SafetyMargin | Should Be 10
        $s.ExpandNestedZips | Should Be $true
        $s.CreateZip | Should Be $false
        $s.AbbreviationsCsv | Should Be (Join-Path $toolRoot 'Abbreviations.csv')
        @($s.RecentDestinations).Count | Should Be 0
    }

    It 'saves settings and reads them back' {
        $s = Read-AppSettings $settingsFile $toolRoot
        $s.MaxPathLength = 259
        $s.CreateZip = $true
        $s.OutputFolder = 'D:\Work\Out'
        Add-RecentDestination $s 'C:\Users\jsmith\CL\Matters - Documents\Jones'
        Save-AppSettings $s $settingsFile
        $r = Read-AppSettings $settingsFile $toolRoot
        $r.MaxPathLength | Should Be 259
        $r.CreateZip | Should Be $true
        $r.OutputFolder | Should Be 'D:\Work\Out'
        @($r.RecentDestinations).Count | Should Be 1
        @($r.RecentDestinations)[0] | Should Be 'C:\Users\jsmith\CL\Matters - Documents\Jones'
    }

    It 'falls back to the defaults if the settings file is damaged' {
        [System.IO.File]::WriteAllText($settingsFile, 'this is { not json')
        $r = Read-AppSettings $settingsFile $toolRoot
        $r.MaxPathLength | Should Be 218
    }

    It 'ignores numbers that are out of range' {
        [System.IO.File]::WriteAllText($settingsFile, '{ "MaxPathLength": 5, "SafetyMargin": 500 }')
        $r = Read-AppSettings $settingsFile $toolRoot
        $r.MaxPathLength | Should Be 218
        $r.SafetyMargin | Should Be 10
    }

    It 'keeps recent SharePoint folders newest first, without duplicates, at most 10' {
        $s = Get-DefaultAppSettings $toolRoot
        Add-RecentDestination $s 'C:\A'
        Add-RecentDestination $s 'C:\B'
        Add-RecentDestination $s 'c:\a\'
        @($s.RecentDestinations) -join '|' | Should Be 'c:\a\|C:\B'
        foreach ($i in 1..15) { Add-RecentDestination $s ('C:\M' + $i) }
        @($s.RecentDestinations).Count | Should Be 10
        @($s.RecentDestinations)[0] | Should Be 'C:\M15'
    }
}

Describe 'Results table and filter' {
    $rows = @(
        (New-TestRow 'Letters\Letter.pdf' 'Letters\Letter.pdf' 'OK'),
        (New-TestRow 'Deep\Very long name.pdf' 'Deep\Very long name.pdf' 'Needs manual attention' 'Folder name shortened'),
        (New-TestRow 'Confidential\Statement.pdf' '' 'Skipped - encrypted' 'Password protected entry'),
        (New-TestRow 'Invoices\Invoice [draft] 100% O''Brien.pdf' 'Invs\Invoice [draft] 100% O''Brien.pdf' 'Renamed' 'Abbreviations')
    )
    $t = ConvertTo-ResultTable $rows $prefix

    It 'builds one table row per report row' {
        ($t -is [System.Data.DataTable]) | Should Be $true
        $t.Rows.Count | Should Be 4
    }

    It 'shows paths relative to the SharePoint folder' {
        $t.Rows[0]['Original'] | Should Be 'Letters\Letter.pdf'
        $t.Rows[3]['New'] | Should Be 'Invs\Invoice [draft] 100% O''Brien.pdf'
        $t.Rows[0]['NewFull'] | Should Be ($prefix + '\Letters\Letter.pdf')
    }

    It 'leaves the length blank for files that are not copied' {
        [System.DBNull]::Value.Equals($t.Rows[2]['NewLength']) | Should Be $true
        [int]$t.Rows[0]['NewLength'] | Should Be ($prefix + '\Letters\Letter.pdf').Length
    }

    It 'filters to problems only' {
        $t.DefaultView.RowFilter = Get-ResultFilter $true ''
        $t.DefaultView.Count | Should Be 2
        $t.DefaultView.RowFilter = ''
    }

    It 'searches safely for names with brackets, percent signs, apostrophes and stars' {
        foreach ($term in @('[draft]', '100%', 'O''Brien', 'invoice [')) {
            $t.DefaultView.RowFilter = Get-ResultFilter $false $term
            $t.DefaultView.Count | Should Be 1
        }
        $t.DefaultView.RowFilter = Get-ResultFilter $false '*'
        $t.DefaultView.Count | Should Be 0
        $t.DefaultView.RowFilter = Get-ResultFilter $true 'letter'
        $t.DefaultView.Count | Should Be 0
        $t.DefaultView.RowFilter = ''
    }

    It 'colours problems red, skipped items amber and unchanged files white' {
        (Get-StatusRgb 'Needs manual attention') -join ',' | Should Be '255,214,214'
        (Get-StatusRgb 'Skipped - encrypted') -join ',' | Should Be '255,242,204'
        (Get-StatusRgb 'OK') -join ',' | Should Be '255,255,255'
    }
}

Describe 'Summary text' {
    It 'describes a check with problems' {
        $res = [pscustomobject]@{
            Mode = 'DryRun'; FilesProcessed = 10; FilesRenamed = 3; FilesFlagged = 1; ItemsSkipped = 2; ItemsRejected = 0; Errors = 0
            Rows = @(1..6 | ForEach-Object { [pscustomobject]@{ Status = 'OK' } })
        }
        $s = Get-ResultSummary $res
        $s.Text | Should Be 'Checked 10 files: 3 will be renamed, 6 already fit, 1 still too long (need attention), 2 skipped.'
        $s.Level | Should Be 'Warning'
        $s.Next | Should Match 'Nothing has been copied yet'
    }

    It 'describes a clean finished apply' {
        $res = [pscustomobject]@{
            Mode = 'Apply'; FilesProcessed = 5; FilesRenamed = 2; FilesFlagged = 0; ItemsSkipped = 0; ItemsRejected = 0; Errors = 0
            Rows = @()
        }
        $s = Get-ResultSummary $res
        $s.Text | Should Be 'Done. 5 files processed: 2 renamed.'
        $s.Level | Should Be 'Good'
    }
}

Describe 'Running in the background' {
    It 'runs a check in the background, reports progress and returns the result' {
        $out = $workRoot + $sep + 'out-bg'
        $job = Start-ShortenerJob -ModulePath $enginePath -Parameters @{
            Source = $fx.ZipPath; DestinationPrefix = $prefix; OutputFolder = $out; ExpandNestedZips = $true; AllowSyncedOutput = $true
        }
        Wait-ForJob $job
        $job.Sync.Finished | Should Be $true
        $job.Sync.Error | Should BeNullOrEmpty
        $job.Sync.Result.Mode | Should Be 'DryRun'
        $job.Sync.Result.FilesProcessed | Should BeGreaterThan 20
        $job.Sync.Message | Should Be 'Writing the report...'
        [System.IO.File]::Exists($job.Sync.Result.ReportPath) | Should Be $true
    }

    It 'stops when Stop is pressed and copies nothing' {
        $out = $workRoot + $sep + 'out-stop'
        $job = Start-ShortenerJob -ModulePath $enginePath -Parameters @{
            Source = $fx.ZipPath; DestinationPrefix = $prefix; OutputFolder = $out; Apply = $true; AllowSyncedOutput = $true
        }
        $job.Sync.Cancel = $true
        Wait-ForJob $job
        $job.Sync.Error | Should Match '^Stopped: cancelled'
        $job.Sync.Result | Should BeNullOrEmpty
        [System.IO.Directory]::Exists($out + $sep + 'Source') | Should Be $false
    }

    It 'stops part way through copying, removes the half-written file and still writes the log' {
        $out = $workRoot + $sep + 'out-stop-mid'
        # A progress table that asks to stop once copying has started.
        $progress = [hashtable]::Synchronized(@{ Cancel = $false; Message = ''; Done = -1; Total = -1 })
        $watcher = [System.Management.Automation.PowerShell]::Create()
        [void]$watcher.AddScript({
                param($p)
                # Checks continuously (no sleeping) so it reacts within the third file.
                $watch = [System.Diagnostics.Stopwatch]::StartNew()
                while ($watch.Elapsed.TotalSeconds -lt 120) {
                    if ([string]$p['Message'] -like 'Copying files (3 of*') { $p['Cancel'] = $true; break }
                }
            }.ToString()).AddArgument($progress)
        $handle = $watcher.BeginInvoke()
        $msg = Get-ErrorMessage { Invoke-PathShortener -Source $fx.ZipPath -DestinationPrefix $prefix -OutputFolder $out -Apply -AllowSyncedOutput -Quiet -Progress $progress }
        [void]$watcher.EndInvoke($handle)
        $watcher.Dispose()
        $msg | Should Match '^Stopped: cancelled'
        $logs = @([System.IO.Directory]::GetFiles($out, '* - Log *.csv'))
        $logs.Count | Should Be 1
        $copied = @(Get-ChildItem -LiteralPath ($out + $sep + 'Source') -Recurse -File)
        $copied.Count | Should BeLessThan 20
        $logRows = @(Import-Csv -LiteralPath $logs[0] | Where-Object { $_.Result -eq 'Copied' })
        $logRows.Count | Should Be $copied.Count
    }
}

Describe 'Synced folder check' {
    $saved = $env:OneDrive
    $fake = $workRoot + $sep + 'OneDrive - Test'
    [void][System.IO.Directory]::CreateDirectory($fake)
    $env:OneDrive = $fake
    try {
        It 'finds an output folder inside OneDrive' {
            Get-SyncedFolderMatch -OutputFolder ($fake + $sep + 'Out') -DestinationPrefix $prefix | Should Be $fake
        }
        It 'returns nothing for an ordinary folder' {
            Get-SyncedFolderMatch -OutputFolder ($workRoot + $sep + 'out-plain') -DestinationPrefix $prefix | Should Be ''
        }
        It 'treats the SharePoint folder itself as synced' -Skip:(-not $isWin) {
            Get-SyncedFolderMatch -OutputFolder ($workRoot + $sep + 'x' + $sep + 'Out') -DestinationPrefix ($workRoot + $sep + 'x') | Should Be ($workRoot + $sep + 'x')
        }
    } finally {
        $env:OneDrive = $saved
    }
}

Describe 'App scripts' {
    $scripts = @(Get-ChildItem -LiteralPath $toolRoot -Recurse -Include '*.ps1', '*.psm1', '*.bat', '*.csv', '*.md')

    It 'every script parses without errors' {
        foreach ($f in @($scripts | Where-Object { $_.Extension -eq '.ps1' -or $_.Extension -eq '.psm1' })) {
            $errors = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errors)
            ('{0}: {1}' -f $f.Name, @($errors).Count) | Should Be ('{0}: 0' -f $f.Name)
        }
    }

    It 'every file is plain ASCII, so Windows PowerShell 5.1 reads it correctly' {
        foreach ($f in $scripts) {
            $bad = @([System.IO.File]::ReadAllBytes($f.FullName) | Where-Object { $_ -gt 127 }).Count
            ('{0}: {1}' -f $f.Name, $bad) | Should Be ('{0}: 0' -f $f.Name)
        }
    }

    foreach ($name in 'Start-App.ps1', 'Start-Interactive.ps1', 'New-DesktopShortcut.ps1') {
        $path = Join-Path $engineDir $name
        $tokens = $null
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)

        It "$name only calls commands that exist" {
            $defined = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
            $used = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ } | Sort-Object -Unique)
            $missing = @($used | Where-Object { $defined -notcontains $_ -and -not (Get-Command $_ -ErrorAction SilentlyContinue) })
            ($missing -join ', ') | Should Be ''
        }

        It "$name only uses variables that are set somewhere" {
            $assigned = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($a in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)) {
                $left = $a.Left
                if ($left -is [System.Management.Automation.Language.ConvertExpressionAst]) { $left = $left.Child }
                if ($left -is [System.Management.Automation.Language.VariableExpressionAst]) { [void]$assigned.Add(($left.VariablePath.UserPath -replace '^(script|global|local):', '')) }
            }
            foreach ($p in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)) { [void]$assigned.Add($p.Name.VariablePath.UserPath) }
            foreach ($f in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)) { [void]$assigned.Add($f.Variable.VariablePath.UserPath) }
            $automatic = @('_', 'true', 'false', 'null', 'PSScriptRoot', 'args', 'this', 'ErrorActionPreference', 'Matches', 'PSHOME')
            $used = @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true) |
                    ForEach-Object { $_.VariablePath.UserPath -replace '^(script|global|local):', '' } | Where-Object { $_ -notmatch '^env:' } | Sort-Object -Unique)
            $unset = @($used | Where-Object { -not $assigned.Contains($_) -and $automatic -notcontains $_ })
            ($unset -join ', ') | Should Be ''
        }
    }
}

# ---------------------------------------------------------------------------

if (-not $env:LPS_KEEP_TEST_FILES) {
    try { [System.IO.Directory]::Delete($workRoot, $true) } catch { Write-Warning "Could not remove $workRoot : $($_.Exception.Message)" }
} else {
    Write-Host "Test files kept in $workRoot"
}
