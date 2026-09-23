<#
    LongPathShortener.App.psm1

    Support code for the app window (Start-App.ps1) that does not draw anything,
    so it can be tested without a screen:
      * reading and saving the user's settings
      * turning report rows into a table for the results grid
      * the search and "problems only" filter
      * the plain-English summary
      * running the engine in the background so the window stays responsive

    Written for Windows PowerShell 5.1. Keep this file plain ASCII.
#>

Set-StrictMode -Version 2.0

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

function Get-AppSettingsPath {
    # Settings are stored per user on this computer only.
    if ($env:LOCALAPPDATA) { $base = $env:LOCALAPPDATA } else { $base = [System.IO.Path]::GetTempPath() }
    return (Join-Path (Join-Path $base 'LongPathShortener') 'settings.json')
}

function Get-DefaultAppSettings {
    param([string]$ToolRoot)
    return [pscustomobject]@{
        OutputFolder       = 'C:\CL\Out'
        MaxPathLength      = 218
        SafetyMargin       = 10
        ExpandNestedZips   = $true
        CreateZip          = $false
        IncludeSystemFiles = $false
        AbbreviationsCsv   = (Join-Path $ToolRoot 'Abbreviations.csv')
        RecentDestinations = [string[]]@()
    }
}

function Read-AppSettings {
    # Starts from the defaults and takes any valid values from the saved file.
    # A missing or damaged file just means the defaults are used.
    param([string]$Path, [string]$ToolRoot)
    $s = Get-DefaultAppSettings $ToolRoot
    $saved = $null
    try {
        if ($Path -and (Test-Path -LiteralPath $Path)) {
            $saved = [System.IO.File]::ReadAllText($Path) | ConvertFrom-Json
        }
    } catch { $saved = $null }
    if ($null -ne $saved) {
        $names = @($saved.PSObject.Properties | ForEach-Object { $_.Name })
        try { if ($names -contains 'OutputFolder' -and $saved.OutputFolder) { $s.OutputFolder = [string]$saved.OutputFolder } } catch { }
        try { if ($names -contains 'MaxPathLength') { $v = [int]$saved.MaxPathLength; if ($v -ge 50 -and $v -le 400) { $s.MaxPathLength = $v } } } catch { }
        try { if ($names -contains 'SafetyMargin') { $v = [int]$saved.SafetyMargin; if ($v -ge 0 -and $v -le 100) { $s.SafetyMargin = $v } } } catch { }
        try { if ($names -contains 'ExpandNestedZips') { $s.ExpandNestedZips = [bool]$saved.ExpandNestedZips } } catch { }
        try { if ($names -contains 'CreateZip') { $s.CreateZip = [bool]$saved.CreateZip } } catch { }
        try { if ($names -contains 'IncludeSystemFiles') { $s.IncludeSystemFiles = [bool]$saved.IncludeSystemFiles } } catch { }
        try { if ($names -contains 'AbbreviationsCsv' -and $null -ne $saved.AbbreviationsCsv) { $s.AbbreviationsCsv = [string]$saved.AbbreviationsCsv } } catch { }
        try {
            if ($names -contains 'RecentDestinations' -and $null -ne $saved.RecentDestinations) {
                $s.RecentDestinations = [string[]]@($saved.RecentDestinations | Where-Object { $_ } | ForEach-Object { [string]$_ })
            }
        } catch { }
    }
    return $s
}

function Save-AppSettings {
    param($Settings, [string]$Path)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    $json = $Settings | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding -ArgumentList $false))
}

function Add-RecentDestination {
    # Most recent first, no duplicates (ignoring capitals), at most $Max.
    param($Settings, [string]$Destination, [int]$Max = 10)
    $d = $Destination.Trim().Trim('"').Trim()
    if (-not $d) { return }
    $list = New-Object System.Collections.Generic.List[string]
    $list.Add($d)
    foreach ($r in @($Settings.RecentDestinations)) {
        if (-not $r) { continue }
        if ([string]::Equals($r.TrimEnd('\'), $d.TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if ($list.Count -ge $Max) { break }
        $list.Add($r)
    }
    $Settings.RecentDestinations = $list.ToArray()
}

# ---------------------------------------------------------------------------
# Results table
# ---------------------------------------------------------------------------

function Get-StatusRank {
    # Lower numbers are listed first. Below 4 counts as a problem.
    param([string]$Status)
    if ($Status -eq 'Needs manual attention') { return 0 }
    if ($Status -like 'Error*') { return 1 }
    if ($Status -like 'Rejected*') { return 2 }
    if ($Status -like 'Skipped*') { return 3 }
    if ($Status -eq 'Renamed') { return 4 }
    if ($Status -like 'Expanded*') { return 5 }
    return 6
}

function Get-StatusRgb {
    # Row colour for each status, as red, green, blue.
    param([string]$Status)
    switch (Get-StatusRank $Status) {
        0 { return @(255, 214, 214) }
        1 { return @(255, 199, 199) }
        2 { return @(255, 214, 214) }
        3 { return @(255, 242, 204) }
        4 { return @(226, 239, 255) }
        5 { return @(236, 234, 250) }
        default { return @(255, 255, 255) }
    }
}

function Get-DisplayPath {
    # Shows a path relative to the SharePoint folder, which is the part that differs.
    param([string]$FullPath, [string]$Prefix)
    if (-not $FullPath) { return '' }
    $p = $Prefix.TrimEnd('\') + '\'
    if ($FullPath.StartsWith($p, [System.StringComparison]::OrdinalIgnoreCase)) { return $FullPath.Substring($p.Length) }
    return $FullPath
}

function ConvertTo-ResultTable {
    param($Rows, [string]$Prefix)
    $t = New-Object System.Data.DataTable -ArgumentList 'Results'
    [void]$t.Columns.Add('Rank', [int])
    [void]$t.Columns.Add('Status', [string])
    [void]$t.Columns.Add('Original', [string])
    [void]$t.Columns.Add('New', [string])
    [void]$t.Columns.Add('NewLength', [int])
    [void]$t.Columns.Add('OriginalLength', [int])
    [void]$t.Columns.Add('Changes', [string])
    [void]$t.Columns.Add('OriginalFull', [string])
    [void]$t.Columns.Add('NewFull', [string])
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $row = $t.NewRow()
        $row['Rank'] = Get-StatusRank ([string]$r.Status)
        $row['Status'] = [string]$r.Status
        $row['Original'] = Get-DisplayPath ([string]$r.OriginalPath) $Prefix
        $row['New'] = Get-DisplayPath ([string]$r.NewPath) $Prefix
        if ([string]$r.NewLength -ne '') { $row['NewLength'] = [int]$r.NewLength } else { $row['NewLength'] = [System.DBNull]::Value }
        if ([string]$r.OriginalLength -ne '') { $row['OriginalLength'] = [int]$r.OriginalLength } else { $row['OriginalLength'] = [System.DBNull]::Value }
        $row['Changes'] = [string]$r.RulesApplied
        $row['OriginalFull'] = [string]$r.OriginalPath
        $row['NewFull'] = [string]$r.NewPath
        $t.Rows.Add($row)
    }
    return , $t
}

function ConvertTo-RowFilterLiteral {
    # Escapes text for a DataView LIKE filter. * % [ ] are wildcards there, so
    # they are wrapped in brackets, and quotes are doubled.
    param([string]$Text)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        if ('*%[]'.IndexOf($ch) -ge 0) { [void]$sb.Append('[').Append($ch).Append(']') }
        elseif ($ch -eq "'") { [void]$sb.Append("''") }
        else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Get-ResultFilter {
    param([bool]$ProblemsOnly, [string]$Search)
    $parts = New-Object System.Collections.Generic.List[string]
    if ($ProblemsOnly) { $parts.Add('Rank < 4') }
    if ($Search -and $Search.Trim()) {
        $e = ConvertTo-RowFilterLiteral $Search.Trim()
        $parts.Add(("(Original LIKE '%{0}%' OR New LIKE '%{0}%' OR Status LIKE '%{0}%' OR Changes LIKE '%{0}%')" -f $e))
    }
    return ($parts.ToArray() -join ' AND ')
}

function Get-ResultSummary {
    # A short plain-English summary of a check or an apply, plus how serious it is.
    param($Result)
    $rows = @($Result.Rows)
    $ok = @($rows | Where-Object { $_.Status -eq 'OK' }).Count
    $renamed = [int]$Result.FilesRenamed
    $flagged = [int]$Result.FilesFlagged
    $skipped = [int]$Result.ItemsSkipped
    $rejected = [int]$Result.ItemsRejected
    $errors = [int]$Result.Errors
    $files = [int]$Result.FilesProcessed

    $level = 'Good'
    if ($flagged -gt 0 -or $skipped -gt 0 -or $rejected -gt 0) { $level = 'Warning' }
    if ($errors -gt 0) { $level = 'Error' }

    $extra = New-Object System.Collections.Generic.List[string]
    if ($flagged -gt 0) { $extra.Add(('{0} still too long (need attention)' -f $flagged)) }
    if ($skipped -gt 0) { $extra.Add(('{0} skipped' -f $skipped)) }
    if ($rejected -gt 0) { $extra.Add(('{0} rejected as unsafe' -f $rejected)) }
    if ($errors -gt 0) { $extra.Add(('{0} could not be copied' -f $errors)) }
    $tail = ''
    if ($extra.Count -gt 0) { $tail = ', ' + ($extra.ToArray() -join ', ') }

    if ($Result.Mode -eq 'Apply') {
        $text = 'Done. {0} files processed: {1} renamed{2}.' -f $files, $renamed, $tail
        $next = 'Upload the contents of the output folder to SharePoint. Keep the evidence log with the matter.'
        if ($flagged -gt 0) { $next = 'Upload the contents of the output folder to SharePoint. Sort out the files in the Needs attention folder by hand.' }
    } else {
        if ($files -eq 0) {
            $text = 'No files were found to copy.'
            $next = 'Check that you chose the right zip file or folder.'
        } else {
            $text = 'Checked {0} files: {1} will be renamed, {2} already fit{3}.' -f $files, $renamed, $ok, $tail
            $next = 'Nothing has been copied yet. Look through the list, then click Apply changes.'
            if ($flagged -gt 0) { $next = 'Nothing has been copied yet. Files still too long will go to a separate Needs attention folder. When ready, click Apply changes.' }
        }
    }
    return [pscustomobject]@{ Text = $text; Next = $next; Level = $level }
}

# ---------------------------------------------------------------------------
# Running the engine in the background
# ---------------------------------------------------------------------------

function Start-ShortenerJob {
    # Runs Invoke-PathShortener on a separate thread. The returned Sync table
    # shows progress (Message, Done, Total), takes Cancel = $true to stop, and
    # gets Result or Error and Finished = $true at the end.
    param([hashtable]$Parameters, [string]$ModulePath)
    $sync = [hashtable]::Synchronized(@{
            Message  = 'Starting...'
            Done     = [long]-1
            Total    = [long]-1
            Cancel   = $false
            Result   = $null
            Error    = ''
            Finished = $false
        })
    $params = @{}
    foreach ($k in $Parameters.Keys) { $params[$k] = $Parameters[$k] }
    $params['Progress'] = $sync
    $params['Quiet'] = $true

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    # The window was started with -ExecutionPolicy Bypass. Give the worker the
    # same, so it can load the engine module. (Windows only. Other systems have
    # no execution policy.)
    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        try { $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass } catch { }
    }
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $rs.Open()
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    $script = {
        param($ModulePath, $Params, $Sync)
        try {
            Import-Module $ModulePath -Force
            $Sync.Result = Invoke-PathShortener @Params
        } catch {
            $Sync.Error = $_.Exception.Message
        } finally {
            $Sync.Finished = $true
        }
    }
    [void]$ps.AddScript($script.ToString()).AddArgument($ModulePath).AddArgument($params).AddArgument($sync)
    $handle = $ps.BeginInvoke()
    return [pscustomobject]@{ Sync = $sync; PowerShell = $ps; Runspace = $rs; Handle = $handle }
}

function Complete-ShortenerJob {
    param($Job)
    if ($null -eq $Job) { return }
    try { [void]$Job.PowerShell.EndInvoke($Job.Handle) } catch { }
    try { $Job.PowerShell.Dispose() } catch { }
    try { $Job.Runspace.Dispose() } catch { }
}

Export-ModuleMember -Function Get-AppSettingsPath, Get-DefaultAppSettings, Read-AppSettings, Save-AppSettings,
    Add-RecentDestination, Get-StatusRank, Get-StatusRgb, Get-DisplayPath, ConvertTo-ResultTable,
    ConvertTo-RowFilterLiteral, Get-ResultFilter, Get-ResultSummary, Start-ShortenerJob, Complete-ShortenerJob
