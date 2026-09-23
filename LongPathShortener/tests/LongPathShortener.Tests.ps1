<#
    LongPathShortener.Tests.ps1

    Pester 3.4 tests (the version built into Windows PowerShell 5.1). Uses only
    synthetic data made by New-TestFixtures.ps1. Run them with Run-Tests.ps1.

    Pester 3 syntax is used on purpose: "Should Be" rather than "Should -Be", and
    setup code sits directly inside each Describe block.
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$toolRoot = Split-Path -Parent $here
Import-Module (Join-Path (Join-Path $toolRoot 'engine') 'LongPathShortener.psm1') -Force
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$isWin = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
$sep = [string][System.IO.Path]::DirectorySeparatorChar

# ---------------------------------------------------------------------------
# Test helpers (independent of the module's own code)
# ---------------------------------------------------------------------------

function ConvertTo-TestLongPath {
    param([string]$P)
    if (-not $isWin) { return $P }
    if ($P.StartsWith('\\?\')) { return $P }
    if ($P.StartsWith('\\')) { return '\\?\UNC\' + $P.Substring(2) }
    return '\\?\' + $P
}

function Get-BytesSha256 {
    param([byte[]]$Bytes)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '') } finally { $sha.Dispose() }
}

function Get-TestFileSha256 {
    param([string]$Path)
    $fs = [System.IO.File]::Open((ConvertTo-TestLongPath $Path), 'Open', 'Read', 'Read')
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '') } finally { $sha.Dispose(); $fs.Dispose() }
}

function Get-TestTree {
    # Every file and folder under a root: relative path (with \), full long path.
    # Callers wrap the result in @() so an empty or one-item result is still a list.
    param([string]$Root)
    $result = New-Object System.Collections.Generic.List[object]
    $rootLong = ConvertTo-TestLongPath $Root
    if (-not [System.IO.Directory]::Exists($rootLong)) { return }
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push(@((New-Object System.IO.DirectoryInfo -ArgumentList $rootLong), ''))
    while ($stack.Count -gt 0) {
        $pair = $stack.Pop()
        foreach ($c in $pair[0].EnumerateFileSystemInfos()) {
            if ($pair[1]) { $rel = $pair[1] + '\' + $c.Name } else { $rel = $c.Name }
            $isDir = ($c -is [System.IO.DirectoryInfo])
            $result.Add([pscustomobject]@{ Rel = $rel; Full = $c.FullName; IsDir = $isDir; Info = $c })
            if ($isDir) { $stack.Push(@($c, $rel)) }
        }
    }
    return $result.ToArray()
}

function Get-TreeSnapshot {
    param([string]$Root)
    $lines = foreach ($e in (Get-TestTree $Root)) {
        if ($e.IsDir) { 'D|' + $e.Rel }
        else { 'F|' + $e.Rel + '|' + $e.Info.Length + '|' + $e.Info.LastWriteTimeUtc.Ticks + '|' + (Get-TestFileSha256 $e.Full) }
    }
    return ((@($lines) | Sort-Object) -join "`n")
}

function Add-ZipEntryInfo {
    # Maps "folder\inner.zip\file" style relative paths to the hash and date of
    # every entry, looking inside zips within zips. Unsafe names are ignored.
    param([byte[]]$Bytes, [string]$BaseRel, $Map)
    $ms = New-Object System.IO.MemoryStream -ArgumentList (, $Bytes)
    $za = New-Object System.IO.Compression.ZipArchive -ArgumentList @($ms, [System.IO.Compression.ZipArchiveMode]::Read, $false, [System.Text.Encoding]::UTF8)
    if (@($za.Entries | Where-Object { $_.FullName.IndexOf([char]0xFFFD) -ge 0 }).Count -gt 0) {
        # Not valid UTF-8, so the names use the old DOS code page.
        $za.Dispose()
        $oem = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
        $ms = New-Object System.IO.MemoryStream -ArgumentList (, $Bytes)
        $za = New-Object System.IO.Compression.ZipArchive -ArgumentList @($ms, [System.IO.Compression.ZipArchiveMode]::Read, $false, $oem)
    }
    try {
        foreach ($e in $za.Entries) {
            $name = $e.FullName
            if ($name.EndsWith('/') -or $name.EndsWith('\')) { continue }
            if ($name -match '^[\\/]' -or $name -match '^[A-Za-z]:[\\/]') { continue }
            $segs = @($name.Split([char[]]@('/', '\')) | Where-Object { $_ -ne '' -and $_ -ne '.' })
            if ($segs -contains '..') { continue }
            $rel = (@(@($BaseRel) + $segs) | Where-Object { $_ }) -join '\'
            try {
                $s = $e.Open()
                $buf = New-Object System.IO.MemoryStream
                $s.CopyTo($buf)
                $s.Dispose()
            } catch { continue }
            $data = $buf.ToArray()
            $Map[$rel] = [pscustomobject]@{ Hash = (Get-BytesSha256 $data); Time = $e.LastWriteTime.DateTime }
            if ($rel -match '\.zip$') { try { Add-ZipEntryInfo $data $rel $Map } catch { } }
        }
    } finally {
        $za.Dispose()
    }
}

function New-OrdinalMap {
    return (New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([System.StringComparer]::Ordinal))
}

function Get-ErrorMessage {
    param([scriptblock]$Script)
    try { & $Script | Out-Null; return '' } catch { return $_.Exception.Message }
}

function Get-LeafName {
    param([string]$Rel)
    $i = $Rel.LastIndexOf('\')
    if ($i -lt 0) { return $Rel }
    return $Rel.Substring($i + 1)
}

function Get-TestExtension {
    param([string]$Name)
    $m = [regex]::Match($Name, '\.[A-Za-z0-9]{1,10}$')
    if ($m.Success -and $m.Index -gt 0) { return $m.Value.ToLowerInvariant() }
    return ''
}

function Join-OutputPath {
    param([string]$Root, [string]$Rel)
    return ($Root + $sep + $Rel.Replace('\', $sep))
}

# ---------------------------------------------------------------------------
# Shared set-up
# ---------------------------------------------------------------------------

$workRoot = [System.IO.Path]::GetTempPath() + 'lps-' + [guid]::NewGuid().ToString('N').Substring(0, 6)
$fx = & (Join-Path $here 'New-TestFixtures.ps1') -Path ($workRoot + $sep + 'fx')
$prefix = 'C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd\'
$prefixNoSlash = $prefix.TrimEnd('\')
$budget = 218 - 10
$abbr = Join-Path $toolRoot 'Abbreviations.csv'
$datePattern = '^(\d{4}-\d{2}-\d{2}|\d{3}) '

# ---------------------------------------------------------------------------
# Unit tests for the naming rules
# ---------------------------------------------------------------------------

InModuleScope LongPathShortener {
    Describe 'Rule 1: clean-up' {
        It 'replaces characters SharePoint rejects' {
            (Get-CleanName 'Re: defects? "urgent" <draft>|final*.pdf').Name | Should Be 'Re_ defects_ _urgent_ _draft__final_.pdf'
        }
        It 'trims spaces and removes trailing periods' {
            (Get-CleanName '  Notes about the meeting. . ').Name | Should Be 'Notes about the meeting'
        }
        It 'collapses repeated spaces' {
            (Get-CleanName 'Draft   with   extra   spaces.docx').Name | Should Be 'Draft with extra spaces.docx'
        }
        It 'renames Windows reserved names, with or without an extension' {
            (Get-CleanName 'CON').Name | Should Be 'CON_'
            (Get-CleanName 'con.txt').Name | Should Be 'con_.txt'
            (Get-CleanName 'LPT1.docx').Name | Should Be 'LPT1_.docx'
            (Get-CleanName 'COM9').Name | Should Be 'COM9_'
            (Get-CleanName 'aux').Name | Should Be 'aux_'
        }
        It 'leaves names that only start like a reserved name alone' {
            (Get-CleanName 'Console notes.txt').Name | Should Be 'Console notes.txt'
            (Get-CleanName 'Auxiliary.pdf').Name | Should Be 'Auxiliary.pdf'
        }
    }

    Describe 'Rules 2 to 4: abbreviations, repeated parent name, filler words' {
        $sample = Join-Path (Split-Path -Parent (Split-Path -Parent (Get-Module LongPathShortener).Path)) 'Abbreviations.csv'
        $abbrevs = @(Import-Abbreviation $sample)

        It 'loads the sample abbreviations file' {
            $abbrevs.Count | Should BeGreaterThan 19
        }
        It 'abbreviates whole words only, ignoring case' {
            Invoke-AbbreviationRule 'CORRESPONDENCE with builder' $abbrevs | Should Be 'Corro with builder'
            Invoke-AbbreviationRule 'Construction Contract' $abbrevs | Should Be 'Constr Contr'
            Invoke-AbbreviationRule 'Reconstructions' $abbrevs | Should Be 'Reconstructions'
        }
        It 'prefers the longest matching phrase' {
            Invoke-AbbreviationRule 'Statement of Claim draft' $abbrevs | Should Be 'SOC draft'
        }
        It 'removes text that repeats the parent folder name' {
            Remove-ParentName 'Smith Pty Ltd - Invoices' @('Smith Pty Ltd') | Should Be 'Invoices'
            Remove-ParentName 'Invoices (Smith Pty Ltd)' @('Smith Pty Ltd') | Should Be 'Invoices'
        }
        It 'does not empty a name that is only the parent name' {
            Remove-ParentName 'Smith Pty Ltd' @('Smith Pty Ltd') | Should BeNullOrEmpty
        }
        It 'removes filler words' {
            Remove-FillerWord 'Minutes of the Meeting and Notes for Owner' | Should Be 'Minutes Meeting Notes Owner'
        }
        It 'keeps words that contain filler words' {
            Remove-FillerWord 'Theatre Fortitude Andrews' | Should BeNullOrEmpty
        }
    }

    Describe 'Rules 5 and 6: truncation' {
        It 'cuts at a word boundary' {
            Get-TruncatedText 'Correspondence with the builder' 20 12 | Should Be 'Correspondence with'
        }
        It 'never cuts a folder below 12 characters' {
            (Get-TruncatedText 'Abcdefghijklmnopqrstuvwxyz' 15 12).Length | Should Be 15
            (Get-TruncatedText 'Level 03 - Variations and claims' 12 12).Length | Should Be 12
        }
        It 'recognises leading dates and document numbers' {
            Get-ProtectedPrefix '2026-03-14 Letter to builder' | Should Be '2026-03-14 '
            Get-ProtectedPrefix '001 Statutory declaration' | Should Be '001 '
            Get-ProtectedPrefix '20260314_Letter' | Should Be '20260314_'
            Get-ProtectedPrefix 'Letter 2026' | Should Be ''
        }
    }

    Describe 'Report safety' {
        It 'stops a client file name from running as an Excel formula' {
            Format-CsvCell '=HYPERLINK("http://x")' | Should Be '"''=HYPERLINK(""http://x"")"'
            Format-CsvCell 'Plain name.pdf' | Should Be '"Plain name.pdf"'
        }
    }
}

# ---------------------------------------------------------------------------
# Zip source, end to end
# ---------------------------------------------------------------------------

Describe 'Zip source' {
    $zipLong = ConvertTo-TestLongPath $fx.ZipPath
    $zipHashBefore = Get-TestFileSha256 $fx.ZipPath
    $zipTimeBefore = [System.IO.File]::GetLastWriteTimeUtc($zipLong)
    $zipLengthBefore = (New-Object System.IO.FileInfo -ArgumentList $zipLong).Length

    $sourceMap = New-OrdinalMap
    Add-ZipEntryInfo ([System.IO.File]::ReadAllBytes($zipLong)) '' $sourceMap

    $out = $workRoot + $sep + 'out-zip'
    $dry = Invoke-PathShortener -Source $fx.ZipPath -OutputFolder $out -DestinationPrefix $prefix -AbbreviationsCsv $abbr -ExpandNestedZips -AllowSyncedOutput -Quiet
    $afterDry = @(Get-TestTree $out)
    $res = Invoke-PathShortener -Source $fx.ZipPath -OutputFolder $out -DestinationPrefix $prefix -AbbreviationsCsv $abbr -ExpandNestedZips -CreateZip -Apply -AllowSyncedOutput -Quiet
    $outTree = @(Get-TestTree $res.FilesFolder)
    $outFiles = @($outTree | Where-Object { -not $_.IsDir })
    $outRels = @($outTree | ForEach-Object { $_.Rel })
    $copied = @($res.LogRows | Where-Object { $_.Result -like 'Copied*' })

    It 'dry run writes the report and nothing else' {
        $afterDry.Count | Should Be 1
        $afterDry[0].Rel | Should Match 'Dry run report .*\.csv$'
        [System.IO.File]::Exists((ConvertTo-TestLongPath $dry.ReportPath)) | Should Be $true
    }

    It 'every output path fits within the limit including the destination prefix' {
        $outTree.Count | Should BeGreaterThan 20
        foreach ($e in $outTree) {
            ($prefixNoSlash + '\' + $e.Rel).Length | Should Not BeGreaterThan $budget
        }
    }

    It 'every renamed or unchanged row in the report fits within the limit' {
        $fitting = @($res.Rows | Where-Object { $_.Status -eq 'OK' -or $_.Status -eq 'Renamed' })
        $fitting.Count | Should BeGreaterThan 15
        foreach ($r in $fitting) { [int]$r.NewLength | Should Not BeGreaterThan $budget }
    }

    It 'flags paths that cannot fit and keeps them out of the main folder' {
        $flagged = @($res.Rows | Where-Object { $_.Status -eq 'Needs manual attention' })
        $flagged.Count | Should BeGreaterThan 0
        foreach ($r in $flagged) { [int]$r.NewLength | Should BeGreaterThan $budget }
        $attention = @(Get-TestTree $res.AttentionFolder | Where-Object { -not $_.IsDir })
        $attention.Count | Should Be $flagged.Count
        @($res.LogRows | Where-Object { $_.Result -eq 'Copied to Needs attention folder' }).Count | Should Be $flagged.Count
    }

    It 'SHA-256 of every output file matches the source entry and the log' {
        $copied.Count | Should BeGreaterThan 20
        foreach ($row in $copied) {
            $sourceMap.ContainsKey($row.OriginalRelativePath) | Should Be $true
            if ($row.Result -eq 'Copied') { $root = $res.FilesFolder } else { $root = $res.AttentionFolder }
            $outHash = Get-TestFileSha256 (Join-OutputPath $root $row.NewRelativePath)
            $outHash | Should Be $sourceMap[$row.OriginalRelativePath].Hash
            $row.SHA256 | Should Be $outHash
        }
    }

    It 'keeps every file extension' {
        foreach ($row in $copied) {
            Get-TestExtension (Get-LeafName $row.NewRelativePath) | Should Be (Get-TestExtension (Get-LeafName $row.OriginalRelativePath))
        }
    }

    It 'keeps leading dates and document numbers' {
        $dated = @($copied | Where-Object { (Get-LeafName $_.OriginalRelativePath) -match $datePattern })
        $dated.Count | Should BeGreaterThan 3
        foreach ($row in $dated) {
            $token = [regex]::Match((Get-LeafName $row.OriginalRelativePath), $datePattern).Value
            (Get-LeafName $row.NewRelativePath).StartsWith($token) | Should Be $true
        }
    }

    It 'keeps each entry''s last-modified time' {
        foreach ($row in @($copied | Where-Object { $_.Result -eq 'Copied' })) {
            $outTime = [System.IO.File]::GetLastWriteTime((ConvertTo-TestLongPath (Join-OutputPath $res.FilesFolder $row.NewRelativePath)))
            [Math]::Abs(($outTime - $sourceMap[$row.OriginalRelativePath].Time).TotalSeconds) | Should Not BeGreaterThan 2
        }
    }

    It 'leaves the original zip untouched' {
        Get-TestFileSha256 $fx.ZipPath | Should Be $zipHashBefore
        [System.IO.File]::GetLastWriteTimeUtc($zipLong) | Should Be $zipTimeBefore
        (New-Object System.IO.FileInfo -ArgumentList $zipLong).Length | Should Be $zipLengthBefore
    }

    It 'rejects every zip slip entry and writes nothing outside the output folder' {
        foreach ($slip in $fx.ZipSlipEntries) {
            $row = @($res.Rows | Where-Object { $_.OriginalPath -eq ($prefixNoSlash + '\' + $slip.Replace('/', '\')) })
            $row.Count | Should Be 1
            $row[0].Status | Should Be 'Rejected - unsafe path (zip slip)'
            $row[0].NewPath | Should BeNullOrEmpty
        }
        @(Get-TestTree $workRoot | Where-Object { $_.Rel -match 'evil-\d\.txt$' }).Count | Should Be 0
        @([System.IO.Directory]::GetFiles((Split-Path -Parent $workRoot), 'evil-*.txt')).Count | Should Be 0
        if ($isWin) { $absolute = @('C:\Windows\evil-4.txt', 'C:\evil-5.txt') } else { $absolute = @('/tmp/evil-3.txt') }
        foreach ($a in $absolute) { [System.IO.File]::Exists($a) | Should Be $false }
        @($res.LogRows | Where-Object { $_.OriginalRelativePath -match 'evil' -and $_.Result -like 'Copied*' }).Count | Should Be 0
    }

    It 'reports and skips the encrypted entry' {
        $row = @($res.Rows | Where-Object { $_.OriginalPath -like '*Encrypted statement.pdf' })
        $row.Count | Should Be 1
        $row[0].Status | Should Be 'Skipped - encrypted'
        @($outRels | Where-Object { $_ -like '*Encrypted statement*' }).Count | Should Be 0
    }

    It 'expands nested zips 3 levels deep and keeps the 4th level as a zip' {
        ($outRels -contains 'Nested\Level1\L1 file.txt') | Should Be $true
        ($outRels -contains 'Nested\Level1\Level2\Level3\L3 file.txt') | Should Be $true
        ($outRels -contains 'Nested\Level1\Level2\Level3\Level4.zip') | Should Be $true
        @($outRels | Where-Object { $_ -like '*L4 file.txt' }).Count | Should Be 0
    }

    It 'numbers duplicate names instead of overwriting them' {
        ($outRels -contains 'Correspondence\Letter.pdf') | Should Be $true
        ($outRels -contains 'Correspondence\LETTER (2).pdf') | Should Be $true
        $letters = @($outFiles | Where-Object { $_.Rel -like 'Letters\*' })
        $letters.Count | Should Be 2
        @($letters | Where-Object { $_.Rel -like '* (2).pdf' }).Count | Should Be 1
    }

    It 'renames reserved names and illegal characters' {
        ($outRels -contains 'COM1_\notes.txt') | Should Be $true
        ($outRels -contains 'LPT1_.docx') | Should Be $true
        ($outRels -contains 'Correspondence\Re_ defects_ _urgent_ _draft__final_.pdf') | Should Be $true
        ($outRels -contains 'Old notes\file.txt') | Should Be $true
    }

    It 'reads accented names from Windows and Mac zips correctly' {
        ($outRels -contains ('Correspondence\Caf' + [char]0xE9 + ' notes.txt')) | Should Be $true
        ($outRels -contains ('Nested\Level1\M' + [char]0xFC + 'ller statement.txt')) | Should Be $true
    }

    It 'skips Mac and Windows housekeeping files' {
        @($outRels | Where-Object { $_ -match '__MACOSX|\.DS_Store$|Thumbs\.db$' }).Count | Should Be 0
    }

    It 'creates a new zip holding the same files' {
        $zipMap = New-OrdinalMap
        Add-ZipEntryInfo ([System.IO.File]::ReadAllBytes((ConvertTo-TestLongPath $res.ZipPath))) '' $zipMap
        foreach ($f in $outFiles) {
            $zipMap.ContainsKey($f.Rel) | Should Be $true
            $zipMap[$f.Rel].Hash | Should Be (Get-TestFileSha256 $f.Full)
        }
    }

    It 'writes the evidence log with the required columns' {
        $header = [System.IO.File]::ReadAllLines((ConvertTo-TestLongPath $res.LogPath))[0].TrimStart([char]0xFEFF)
        $header | Should Be '"OriginalRelativePath","NewRelativePath","SizeBytes","LastModified","SHA256","Result"'
        $report = [System.IO.File]::ReadAllLines((ConvertTo-TestLongPath $res.ReportPath))[0].TrimStart([char]0xFEFF)
        $report | Should Be '"OriginalPath","NewPath","OriginalLength","NewLength","RulesApplied","Status"'
    }
}

# ---------------------------------------------------------------------------
# Folder source, end to end
# ---------------------------------------------------------------------------

Describe 'Folder source' {
    $before = Get-TreeSnapshot $fx.FolderPath

    $sourceMap = New-OrdinalMap
    foreach ($e in (Get-TestTree $fx.FolderPath)) {
        if ($e.IsDir) { continue }
        $sourceMap[$e.Rel] = [pscustomobject]@{ Hash = (Get-TestFileSha256 $e.Full); Time = $e.Info.LastWriteTimeUtc }
        if ($e.Rel -match '\.zip$') { Add-ZipEntryInfo ([System.IO.File]::ReadAllBytes($e.Full)) $e.Rel $sourceMap }
    }

    $out = $workRoot + $sep + 'out-folder'
    $res = Invoke-PathShortener -Source $fx.FolderPath -OutputFolder $out -DestinationPrefix $prefix -AbbreviationsCsv $abbr -ExpandNestedZips -Apply -AllowSyncedOutput -Quiet
    $after = Get-TreeSnapshot $fx.FolderPath
    $outTree = @(Get-TestTree $res.FilesFolder)
    $outRels = @($outTree | ForEach-Object { $_.Rel })
    $copied = @($res.LogRows | Where-Object { $_.Result -like 'Copied*' })

    It 'every output path fits within the limit including the destination prefix' {
        $outTree.Count | Should BeGreaterThan 15
        foreach ($e in $outTree) {
            ($prefixNoSlash + '\' + $e.Rel).Length | Should Not BeGreaterThan $budget
        }
    }

    It 'SHA-256 of every output file matches the source file and the log' {
        $copied.Count | Should BeGreaterThan 15
        foreach ($row in $copied) {
            $sourceMap.ContainsKey($row.OriginalRelativePath) | Should Be $true
            if ($row.Result -eq 'Copied') { $root = $res.FilesFolder } else { $root = $res.AttentionFolder }
            $outHash = Get-TestFileSha256 (Join-OutputPath $root $row.NewRelativePath)
            $outHash | Should Be $sourceMap[$row.OriginalRelativePath].Hash
            $row.SHA256 | Should Be $outHash
        }
    }

    It 'keeps every file extension' {
        foreach ($row in $copied) {
            Get-TestExtension (Get-LeafName $row.NewRelativePath) | Should Be (Get-TestExtension (Get-LeafName $row.OriginalRelativePath))
        }
    }

    It 'keeps leading dates and document numbers' {
        $dated = @($copied | Where-Object { (Get-LeafName $_.OriginalRelativePath) -match $datePattern })
        $dated.Count | Should BeGreaterThan 3
        foreach ($row in $dated) {
            $token = [regex]::Match((Get-LeafName $row.OriginalRelativePath), $datePattern).Value
            (Get-LeafName $row.NewRelativePath).StartsWith($token) | Should Be $true
        }
    }

    It 'keeps each file''s last-modified time' {
        foreach ($row in @($copied | Where-Object { $_.Result -eq 'Copied' -and $_.OriginalRelativePath -notmatch '\.zip\\' })) {
            $outTime = [System.IO.File]::GetLastWriteTimeUtc((ConvertTo-TestLongPath (Join-OutputPath $res.FilesFolder $row.NewRelativePath)))
            $outTime | Should Be $sourceMap[$row.OriginalRelativePath].Time
        }
    }

    It 'leaves the source folder untouched' {
        @($before -split "`n").Count | Should BeGreaterThan 20
        $after | Should Be $before
    }

    It 'renames reserved names and keeps empty folders' {
        ($outRels -contains 'PRN_\readme.txt') | Should Be $true
        ($outRels -contains 'Empty folder') | Should Be $true
        ($outRels -contains 'Nested\Inner documents\Inner letter.pdf') | Should Be $true
    }

    It 'refuses an output folder inside the source folder' {
        $msg = Get-ErrorMessage { Invoke-PathShortener -Source $fx.FolderPath -OutputFolder ($fx.FolderPath + $sep + 'Out') -DestinationPrefix $prefix -AllowSyncedOutput -Quiet }
        $msg | Should Match 'overlap'
        [System.IO.Directory]::Exists((ConvertTo-TestLongPath ($fx.FolderPath + $sep + 'Out'))) | Should Be $false
    }
}

# ---------------------------------------------------------------------------
# Safety checks
# ---------------------------------------------------------------------------

Describe 'Safety checks' {
    It 'stops before extracting anything when the zip is over the size limit' {
        $out = $workRoot + $sep + 'out-size'
        $msg = Get-ErrorMessage { Invoke-PathShortener -Source $fx.ZipPath -OutputFolder $out -DestinationPrefix $prefix -MaxTotalSizeGB 0.00001 -Apply -AllowSyncedOutput -Quiet }
        $msg | Should Match 'over the .* GB limit'
        @(Get-TestTree $out | Where-Object { -not $_.IsDir }).Count | Should Be 0
    }

    It 'rejects a destination prefix that is not a full path' {
        $msg = Get-ErrorMessage { Invoke-PathShortener -Source $fx.ZipPath -OutputFolder ($workRoot + $sep + 'out-x') -DestinationPrefix 'Matters\Smith' -AllowSyncedOutput -Quiet }
        $msg | Should Match 'must be a full path'
    }

    $savedOneDrive = $env:OneDrive
    $fakeOneDrive = $workRoot + $sep + 'OneDrive - Test'
    [void][System.IO.Directory]::CreateDirectory($fakeOneDrive)
    $env:OneDrive = $fakeOneDrive
    try {
        It 'asks for confirmation when the output folder is inside OneDrive, and stops on no' {
            Mock -ModuleName LongPathShortener Read-Host { 'no' }
            $msg = Get-ErrorMessage { Invoke-PathShortener -Source $fx.ZipPath -OutputFolder ($fakeOneDrive + $sep + 'Out') -DestinationPrefix $prefix -Quiet 3>$null }
            $msg | Should Match 'Stopped. Choose an output folder outside OneDrive'
            Assert-MockCalled -ModuleName LongPathShortener Read-Host -Times 1 -Exactly
            [System.IO.Directory]::Exists($fakeOneDrive + $sep + 'Out') | Should Be $false
        }

        It 'carries on when the user types YES' {
            Mock -ModuleName LongPathShortener Read-Host { 'YES' }
            $r = Invoke-PathShortener -Source $fx.ZipPath -OutputFolder ($fakeOneDrive + $sep + 'Out') -DestinationPrefix $prefix -Quiet 3>$null
            $r.OutputIsSynced | Should Be $true
            [System.IO.File]::Exists((ConvertTo-TestLongPath $r.ReportPath)) | Should Be $true
        }
    } finally {
        $env:OneDrive = $savedOneDrive
    }
}

# ---------------------------------------------------------------------------
# Clean up (set LPS_KEEP_TEST_FILES=1 to keep the fixtures and output)
# ---------------------------------------------------------------------------

if (-not $env:LPS_KEEP_TEST_FILES) {
    try { [System.IO.Directory]::Delete((ConvertTo-TestLongPath $workRoot), $true) } catch { Write-Warning "Could not remove $workRoot : $($_.Exception.Message)" }
} else {
    Write-Host "Test files kept in $workRoot"
}
