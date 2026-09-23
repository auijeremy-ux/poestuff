<#
    LongPathShortener.psm1

    Shortens folder and file names so that every path fits under a length limit
    once the files sit in their final synced SharePoint folder.

    Ground rules this module never breaks:
      * Works offline. It makes no network calls of any kind.
      * Never writes to, renames or deletes anything in the source folder or zip.
      * Never changes file contents. Each copied file is hashed (SHA-256) from the
        source bytes as they are read, hashed again after it is written, and the
        run stops if the two hashes differ.
      * All file access goes through .NET System.IO with the \\?\ long path prefix,
        so source paths longer than 260 characters can be read.

    Written for Windows PowerShell 5.1. Keep this file plain ASCII so that
    PowerShell 5.1 reads it correctly.
#>

Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------------
# Settings and constants
# ---------------------------------------------------------------------------

$script:IsWindowsOS = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
$script:Sep = [string][System.IO.Path]::DirectorySeparatorChar
$script:Quiet = $false
$script:Progress = $null          # optional shared hashtable used by the app window
$script:CancelMessage = 'Stopped: cancelled by the user.'

$script:MinFolderLength = 12      # rule 5 never cuts a folder name below this
$script:MinFileStemLength = 12    # rule 6 never cuts a file name (without extension) below this
$script:MaxNameLength = 240       # a single name is always kept under the Windows 255 limit
$script:MaxNestedZipLevel = 3     # zips inside zips are expanded this many levels deep

$script:ClutterFileNames = @('.DS_Store', 'Thumbs.db', 'desktop.ini')
$script:ClutterFolderNames = @('__MACOSX')

$script:TrimChars = [char[]]@(' ', '_', '-', ',', '.', '(', '[', '&', '+', ';', [char]0x2013, [char]0x2014)
$script:BoundaryChars = [char[]]@(' ', '_', '-', ',', [char]0x2013, [char]0x2014)

$script:IllegalCharRegex = New-Object System.Text.RegularExpressions.Regex -ArgumentList '[\x00-\x1F"*:<>?/\\|]'
$script:ReservedNameRegex = '^(?:CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$'
$script:ExtensionRegex = '\.[A-Za-z0-9]{1,10}$'

# A leading date (2026-03-14, 20260314, 14.03.2026) or document number (001, 1.2.3)
# plus the separator after it. Rule 6 never cuts into this part of a file name.
$script:ProtectedPrefixRegex = '^(?:\d{4}[-._]\d{1,2}[-._]\d{1,2}|\d{8}|\d{1,2}[-._]\d{1,2}[-._]\d{2,4}|\d+(?:\.\d+)*)(?=$|[\s\-_.,)\]])[\s\-_.,)\]]*'

$script:FillerRegex = New-Object System.Text.RegularExpressions.Regex -ArgumentList @(
    '(?<![\p{L}\p{N}''\u2019\-])(?:the|and|of|for)(?![\p{L}\p{N}''\u2019\-])',
    ([System.Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant')
)

# Labels used in the RulesApplied column, in the order the rules run.
$script:RuleLabelOrder = @(
    'Nested zip expanded',
    'Clean-up (illegal characters)',
    'Clean-up (extra spaces)',
    'Clean-up (leading or trailing spaces or periods)',
    'Clean-up (reserved Windows name)',
    'Clean-up (name SharePoint blocks)',
    'Very long name shortened',
    'Abbreviations',
    'Repeated parent name removed',
    'Filler words removed',
    'Folder name shortened',
    'File name shortened',
    'Duplicate name numbered'
)
$script:DuplicateLabel = 'Duplicate name numbered'

# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

function Write-Info {
    param([string]$Message = '', [string]$Color = '')
    if ($script:Quiet) { return }
    if ($Color) { Write-Host $Message -ForegroundColor $Color } else { Write-Host $Message }
}

function Update-RunProgress {
    # Reports progress to the app window (if there is one) and stops the run
    # when the user has pressed Stop. Does nothing on the command line.
    param([string]$Message = '', [long]$Done = -1, [long]$Total = -1)
    $p = $script:Progress
    if ($null -eq $p) { return }
    if ($p['Cancel']) { throw $script:CancelMessage }
    if ($Message) { $p['Message'] = $Message }
    $p['Done'] = $Done
    $p['Total'] = $Total
}

function Add-RunWarning {
    param($Ctx, [string]$Message)
    $Ctx.Warnings.Add($Message)
    if (-not $script:Quiet) { Write-Warning $Message }
}

function ConvertTo-LongPath {
    # Adds the \\?\ prefix so Windows skips its 260 character limit.
    # With this prefix Windows does no clean-up of the path, so callers must pass
    # a full path with backslashes only and no . or .. parts.
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not $script:IsWindowsOS) { return $Path }
    if ($Path.StartsWith('\\?\')) { return $Path }
    if ($Path.StartsWith('\\')) { return '\\?\UNC\' + $Path.Substring(2) }
    return '\\?\' + $Path
}

function ConvertFrom-LongPath {
    param([string]$Path)
    if ($Path.StartsWith('\\?\UNC\')) { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\')) { return $Path.Substring(4) }
    return $Path
}

function Remove-DotSegment {
    # Removes . and .. parts from a full path without touching the disk.
    param([string]$Path)
    if ($script:IsWindowsOS) {
        if ($Path.StartsWith('\\')) {
            $parts = $Path.Substring(2).Split('\')
            if ($parts.Count -lt 2) { return $Path }
            $root = '\\' + $parts[0] + '\' + $parts[1]
            $start = 2
        } else {
            $parts = $Path.Split('\')
            $root = $parts[0]
            $start = 1
        }
        $sepChar = '\'
    } else {
        $parts = $Path.Split('/')
        $root = ''
        $start = 1
        $sepChar = '/'
    }
    $keep = New-Object System.Collections.Generic.List[string]
    for ($i = $start; $i -lt $parts.Count; $i++) {
        $p = $parts[$i]
        if ($p -eq '' -or $p -eq '.') { continue }
        if ($p -eq '..') {
            if ($keep.Count -gt 0) { $keep.RemoveAt($keep.Count - 1) }
            continue
        }
        $keep.Add($p)
    }
    if ($keep.Count -eq 0) { return $root + $sepChar }
    return $root + $sepChar + ($keep -join $sepChar)
}

function Resolve-FullPath {
    # Turns what the user typed or dragged in into a full path (without the \\?\ prefix).
    param([Parameter(Mandatory = $true)][string]$Path)
    $p = $Path.Trim().Trim('"').Trim()
    if ($script:IsWindowsOS) {
        $p = ConvertFrom-LongPath ($p.Replace('/', '\'))
        if ($p -notmatch '^[A-Za-z]:\\' -and $p -notmatch '^\\\\[^\\]') {
            $p = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
        }
    } else {
        if (-not $p.StartsWith('/')) {
            $p = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
        }
    }
    return (Remove-DotSegment $p)
}

function Test-PathUnder {
    # True when $Child is $Parent or sits somewhere inside it.
    param([string]$Child, [string]$Parent)
    if (-not $Child -or -not $Parent) { return $false }
    $c = (ConvertFrom-LongPath $Child).TrimEnd('\', '/')
    $p = (ConvertFrom-LongPath $Parent).TrimEnd('\', '/')
    if ($p -eq '') { return $false }
    if ([string]::Equals($c, $p, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return ($c.StartsWith($p + '\', [System.StringComparison]::OrdinalIgnoreCase) -or
            $c.StartsWith($p + '/', [System.StringComparison]::OrdinalIgnoreCase))
}

function Join-PhysicalPath {
    param([string]$Root, [string[]]$Names)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($Root.TrimEnd('\', '/'))
    foreach ($n in $Names) { [void]$sb.Append($script:Sep).Append($n) }
    return $sb.ToString()
}

function Get-ParentPhysicalPath {
    param([string]$Path)
    return $Path.Substring(0, $Path.LastIndexOf($script:Sep))
}

function Test-NonEmptyDirectory {
    param([string]$LongPath)
    if ([System.IO.File]::Exists($LongPath)) { return $true }
    if (-not [System.IO.Directory]::Exists($LongPath)) { return $false }
    $e = [System.IO.Directory]::EnumerateFileSystemEntries($LongPath).GetEnumerator()
    try { return $e.MoveNext() } finally { $e.Dispose() }
}

function Format-LocalTime {
    param($Date)
    if ($null -eq $Date) { return '' }
    $d = [datetime]$Date
    if ($d.Kind -eq [System.DateTimeKind]::Utc) { $d = $d.ToLocalTime() }
    return $d.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Copy-StreamWithHash {
    # Copies a stream while hashing the bytes read. MaxBytes -1 means no limit.
    param([System.IO.Stream]$InputStream, [System.IO.Stream]$OutputStream, [long]$MaxBytes = -1)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $buffer = New-Object byte[] 1048576
        $total = [long]0
        while ($true) {
            $n = $InputStream.Read($buffer, 0, $buffer.Length)
            if ($n -le 0) { break }
            if ($null -ne $script:Progress -and $script:Progress['Cancel']) { throw $script:CancelMessage }
            $total += $n
            if ($MaxBytes -ge 0 -and $total -gt $MaxBytes) {
                throw (New-Object System.IO.InvalidDataException -ArgumentList 'The entry holds more data than the zip says it should (possible zip bomb), so it was not extracted.')
            }
            [void]$sha.TransformBlock($buffer, 0, $n, $null, 0)
            if ($null -ne $OutputStream) { $OutputStream.Write($buffer, 0, $n) }
        }
        [void]$sha.TransformFinalBlock($buffer, 0, 0)
        return [pscustomobject]@{
            Bytes = $total
            Hash  = [System.BitConverter]::ToString($sha.Hash).Replace('-', '')
        }
    } finally {
        $sha.Dispose()
    }
}

function Get-FileSha256 {
    param([string]$LongPath)
    $fs = [System.IO.File]::Open($LongPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try { return (Copy-StreamWithHash -InputStream $fs -OutputStream $null).Hash } finally { $fs.Dispose() }
}

function Read-Exact {
    param([System.IO.Stream]$Stream, [byte[]]$Buffer, [int]$Count)
    $read = 0
    while ($read -lt $Count) {
        $n = $Stream.Read($Buffer, $read, $Count - $read)
        if ($n -le 0) { throw 'Unexpected end of file.' }
        $read += $n
    }
}

# ---------------------------------------------------------------------------
# CSV output (UTF-8 with BOM so Excel shows accented names correctly)
# ---------------------------------------------------------------------------

function Format-CsvCell {
    param($Value)
    if ($null -eq $Value) { $s = '' } else { $s = [string]$Value }
    # A client-supplied name starting with = + - or @ would run as a formula in
    # Excel. A leading apostrophe makes Excel treat it as plain text.
    if ($s.Length -gt 0 -and '=+-@'.IndexOf($s[0]) -ge 0) { $s = "'" + $s }
    if ($s.Length -gt 0 -and ($s[0] -eq "`t" -or $s[0] -eq "`r")) { $s = "'" + $s }
    return '"' + $s.Replace('"', '""') + '"'
}

function Export-CsvFile {
    param([string]$LongPath, [string[]]$Columns, $Rows)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',')).Append("`r`n")
    foreach ($r in $Rows) {
        $cells = foreach ($c in $Columns) { Format-CsvCell $r.$c }
        [void]$sb.Append(($cells -join ',')).Append("`r`n")
    }
    [System.IO.File]::WriteAllText($LongPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding -ArgumentList $true))
}

function Read-TextFileSmart {
    # Reads a text file saved by Excel or Notepad. Handles UTF-8 (with or without
    # BOM), UTF-16 and falls back to the Windows ANSI code page.
    param([string]$LongPath)
    $bytes = [System.IO.File]::ReadAllBytes($LongPath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    try {
        $strict = New-Object System.Text.UTF8Encoding -ArgumentList $false, $true
        return $strict.GetString($bytes)
    } catch {
        try { return [System.Text.Encoding]::GetEncoding(1252).GetString($bytes) } catch { return [System.Text.Encoding]::Default.GetString($bytes) }
    }
}

# ---------------------------------------------------------------------------
# Naming rules
# ---------------------------------------------------------------------------

function Get-CleanName {
    # Rule 1: clean-up that always applies. Returns the new name and what changed.
    param([string]$Name)
    $changes = New-Object System.Collections.Generic.List[string]

    $n = $script:IllegalCharRegex.Replace($Name, '_')
    if ($n -ne $Name) { $changes.Add('Clean-up (illegal characters)') }

    $collapsed = [regex]::Replace($n, '[ \u00A0]{2,}', ' ')
    if ($collapsed -ne $n) { $changes.Add('Clean-up (extra spaces)') }
    $n = $collapsed

    $trimmed = $n
    do {
        $before = $trimmed
        $trimmed = $trimmed.Trim().TrimEnd('.')
    } while ($trimmed -ne $before)
    if ($trimmed -ne $n) { $changes.Add('Clean-up (leading or trailing spaces or periods)') }
    $n = $trimmed
    if ($n -eq '') { $n = '_' }

    $blocked = $n
    if ($blocked.StartsWith('~$')) { $blocked = '~_' + $blocked.Substring(2) }
    if ($blocked.IndexOf('_vti_', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $blocked = [regex]::Replace($blocked, '_vti_', '_vti-', 'IgnoreCase')
    }
    if ($blocked -eq '.lock') { $blocked = '_lock' }
    if ($blocked -ne $n) { $changes.Add('Clean-up (name SharePoint blocks)') }
    $n = $blocked

    $dot = $n.IndexOf('.')
    if ($dot -lt 0) { $stem = $n } else { $stem = $n.Substring(0, $dot) }
    if ($stem.TrimEnd() -match $script:ReservedNameRegex) {
        if ($dot -lt 0) { $n = $n + '_' } else { $n = $stem.TrimEnd() + '_' + $n.Substring($dot) }
        $changes.Add('Clean-up (reserved Windows name)')
    }

    return [pscustomobject]@{ Name = $n; Changes = $changes.ToArray() }
}

function Split-FileName {
    # Splits a file name into stem and extension. Only a short alphanumeric ending
    # counts as an extension, so "Notes v2.final draft" keeps its whole name.
    param([string]$Name)
    $m = [regex]::Match($Name, $script:ExtensionRegex)
    if ($m.Success -and $m.Index -gt 0) {
        return [pscustomobject]@{ Stem = $Name.Substring(0, $m.Index); Ext = $m.Value }
    }
    return [pscustomobject]@{ Stem = $Name; Ext = '' }
}

function Get-ProtectedPrefix {
    param([string]$Stem)
    $m = [regex]::Match($Stem, $script:ProtectedPrefixRegex)
    if ($m.Success) { return $m.Value }
    return ''
}

function Get-TruncatedText {
    # Cuts text to at most MaxLength characters, backing off to the last word
    # boundary when there is one at or after MinLength.
    param([string]$Text, [int]$MaxLength, [int]$MinLength = 0)
    if ($Text.Length -le $MaxLength) { return $Text }
    if ($MaxLength -lt 1) { return '' }
    $cut = $Text.Substring(0, $MaxLength)
    $atBoundary = ($script:BoundaryChars -contains $Text[$MaxLength])
    if (-not $atBoundary) {
        $idx = $cut.LastIndexOfAny($script:BoundaryChars)
        if ($idx -gt 0 -and $idx -ge $MinLength) { $cut = $cut.Substring(0, $idx) }
    }
    $result = $cut.TrimEnd($script:TrimChars)
    if ($result.Length -eq 0 -or $result.Length -lt $MinLength) {
        $result = $Text.Substring(0, $MaxLength).TrimEnd(' ', '.')
    }
    return $result
}

function Repair-Separators {
    # Tidies up what is left after words are removed from a name.
    param([string]$Text)
    $t = [regex]::Replace($Text, '\(\s*\)|\[\s*\]', '')
    $t = [regex]::Replace($t, '\s*([-_,\u2013\u2014])(?:\s*[-_,\u2013\u2014])+\s*', [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $s = $m.Groups[1].Value
            if ($s -eq ',') { return ', ' }
            if ($m.Value -match '\s') { return ' ' + $s + ' ' }
            return $s
        })
    $t = [regex]::Replace($t, '\s{2,}', ' ')
    $t = $t.Trim([char[]]@(' ', '-', '_', ',', '.', [char]0x2013, [char]0x2014))
    return $t
}

function Test-HasWordCharacter {
    param([string]$Text)
    return ($Text -match '[\p{L}\p{N}]')
}

function Import-Abbreviation {
    # Loads the abbreviations CSV (columns Find, Replace). Longer phrases are tried
    # first so "Statement of Claim" wins over "Statement".
    param([string]$LongPath)
    $text = Read-TextFileSmart $LongPath
    $rows = @($text | ConvertFrom-Csv)
    if ($rows.Count -eq 0) { return @() }
    $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    if (-not ($names -contains 'Find') -or -not ($names -contains 'Replace')) {
        throw "The abbreviations file must have two columns named Find and Replace."
    }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        $find = ([string]$r.Find).Trim()
        $rep = ([string]$r.Replace).Trim()
        if ($find -eq '') { continue }
        $rep = $script:IllegalCharRegex.Replace($rep, '_')
        $pattern = '(?<![\p{L}\p{N}])' + [regex]::Escape($find) + '(?![\p{L}\p{N}])'
        $regex = New-Object System.Text.RegularExpressions.Regex -ArgumentList @($pattern, ([System.Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant'))
        $list.Add([pscustomobject]@{ Find = $find; Replace = $rep; Regex = $regex; Replacement = $rep.Replace('$', '$$') })
    }
    return @($list | Sort-Object -Property @{ Expression = { $_.Find.Length }; Descending = $true })
}

function Invoke-AbbreviationRule {
    param([string]$Text, $Abbreviations)
    $r = $Text
    foreach ($a in $Abbreviations) { $r = $a.Regex.Replace($r, $a.Replacement) }
    return [regex]::Replace($r, '\s{2,}', ' ').Trim()
}

function Remove-ParentName {
    # Rule 3: removes text that repeats a parent folder's name.
    param([string]$Text, [string[]]$ParentNames)
    $candidates = @($ParentNames | Where-Object { $_ -and $_.Trim().Length -ge 3 } | Sort-Object -Property Length -Descending -Unique)
    foreach ($p in $candidates) {
        $pattern = '(?<![\p{L}\p{N}])' + [regex]::Escape($p.Trim()) + '(?![\p{L}\p{N}])'
        if ([regex]::IsMatch($Text, $pattern, 'IgnoreCase')) {
            $r = Repair-Separators ([regex]::Replace($Text, $pattern, '', 'IgnoreCase'))
            if (Test-HasWordCharacter $r) { return $r }
        }
    }
    return $null
}

function Remove-FillerWord {
    # Rule 4: removes "the", "and", "of", "for" (folders only).
    param([string]$Text)
    $r = Repair-Separators ($script:FillerRegex.Replace($Text, ''))
    if ((Test-HasWordCharacter $r) -and $r.Length -lt $Text.Length) { return $r }
    return $null
}

# ---------------------------------------------------------------------------
# Zip reading
# ---------------------------------------------------------------------------

function Get-ZipNameEncoding {
    # Picks how to read entry names that are not marked as UTF-8. Zips made by
    # Windows "Send to compressed folder" use the old DOS (OEM) code page, while
    # Mac and many other tools use UTF-8 without saying so. If every such name is
    # valid UTF-8 it is read as UTF-8, otherwise as the OEM code page. Entries
    # marked as UTF-8 are always read as UTF-8.
    param($Flags)
    $utf8Strict = New-Object System.Text.UTF8Encoding -ArgumentList $false, $true
    $needsOem = $false
    if ($null -ne $Flags) {
        foreach ($f in $Flags) {
            if (($f.Flags -band 0x800) -ne 0) { continue }
            $hasHigh = $false
            foreach ($b in $f.NameBytes) { if ($b -ge 0x80) { $hasHigh = $true; break } }
            if (-not $hasHigh) { continue }
            try { [void]$utf8Strict.GetString($f.NameBytes) } catch { $needsOem = $true; break }
        }
    }
    if (-not $needsOem) { return [System.Text.Encoding]::UTF8 }
    try {
        $cp = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage
        return [System.Text.Encoding]::GetEncoding($cp)
    } catch {
        try { return [System.Text.Encoding]::GetEncoding(437) } catch { return [System.Text.Encoding]::UTF8 }
    }
}

function Get-ZipEntryFlag {
    # .NET Framework cannot tell whether a zip entry is encrypted, so this reads
    # the zip's own index (the central directory) and returns the flags and
    # compression method of every entry in order. Read only. Returns $null if the
    # index cannot be read.
    param([string]$LongPath)
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($LongPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $len = $fs.Length
        if ($len -lt 22) { return $null }
        $tailLen = [int][Math]::Min($len, 65557)
        $tail = New-Object byte[] $tailLen
        [void]$fs.Seek($len - $tailLen, [System.IO.SeekOrigin]::Begin)
        Read-Exact $fs $tail $tailLen
        $eocd = -1
        for ($i = $tailLen - 22; $i -ge 0; $i--) {
            if ($tail[$i] -eq 0x50 -and $tail[$i + 1] -eq 0x4B -and $tail[$i + 2] -eq 0x05 -and $tail[$i + 3] -eq 0x06) { $eocd = $i; break }
        }
        if ($eocd -lt 0) { return $null }
        $count = [long][System.BitConverter]::ToUInt16($tail, $eocd + 10)
        $cdSize = [long][System.BitConverter]::ToUInt32($tail, $eocd + 12)
        $cdOffset = [long][System.BitConverter]::ToUInt32($tail, $eocd + 16)
        $cdEnd = $len - $tailLen + $eocd
        if ($count -eq 0xFFFF -or $cdSize -eq 0xFFFFFFFF -or $cdOffset -eq 0xFFFFFFFF) {
            $loc = $eocd - 20
            if ($loc -lt 0) { return $null }
            if (-not ($tail[$loc] -eq 0x50 -and $tail[$loc + 1] -eq 0x4B -and $tail[$loc + 2] -eq 0x06 -and $tail[$loc + 3] -eq 0x07)) { return $null }
            $z64Offset = [long][System.BitConverter]::ToUInt64($tail, $loc + 8)
            $z64 = New-Object byte[] 56
            [void]$fs.Seek($z64Offset, [System.IO.SeekOrigin]::Begin)
            Read-Exact $fs $z64 56
            if (-not ($z64[0] -eq 0x50 -and $z64[1] -eq 0x4B -and $z64[2] -eq 0x06 -and $z64[3] -eq 0x06)) { return $null }
            $cdSize = [long][System.BitConverter]::ToUInt64($z64, 40)
            $cdEnd = $z64Offset
        }
        $cdStart = $cdEnd - $cdSize
        if ($cdStart -lt 0 -or $cdSize -gt 512MB) { return $null }
        $cd = New-Object byte[] ([int]$cdSize)
        [void]$fs.Seek($cdStart, [System.IO.SeekOrigin]::Begin)
        Read-Exact $fs $cd ([int]$cdSize)
        $list = New-Object System.Collections.Generic.List[object]
        $p = 0
        while ($p + 46 -le $cdSize) {
            if (-not ($cd[$p] -eq 0x50 -and $cd[$p + 1] -eq 0x4B -and $cd[$p + 2] -eq 0x01 -and $cd[$p + 3] -eq 0x02)) { break }
            $flags = [int][System.BitConverter]::ToUInt16($cd, $p + 8)
            $method = [int][System.BitConverter]::ToUInt16($cd, $p + 10)
            $nameLen = [int][System.BitConverter]::ToUInt16($cd, $p + 28)
            $extraLen = [int][System.BitConverter]::ToUInt16($cd, $p + 30)
            $commentLen = [int][System.BitConverter]::ToUInt16($cd, $p + 32)
            $nameBytes = New-Object byte[] $nameLen
            [System.Array]::Copy($cd, $p + 46, $nameBytes, 0, $nameLen)
            $list.Add([pscustomobject]@{ Flags = $flags; Method = $method; NameBytes = $nameBytes })
            $p += 46 + $nameLen + $extraLen + $commentLen
        }
        return , $list.ToArray()
    } catch {
        return $null
    } finally {
        if ($null -ne $fs) { $fs.Dispose() }
    }
}

function Get-OpenArchive {
    # Opens a zip read-only (and keeps it open for the rest of the run).
    param($Ctx, [string]$LongPath, $NameEncoding = $null)
    if ($Ctx.Archives.ContainsKey($LongPath)) { return $Ctx.Archives[$LongPath] }
    if ($null -eq $NameEncoding) { $NameEncoding = [System.Text.Encoding]::UTF8 }
    $fs = [System.IO.File]::Open($LongPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $za = New-Object System.IO.Compression.ZipArchive -ArgumentList @($fs, [System.IO.Compression.ZipArchiveMode]::Read, $false, $NameEncoding)
    } catch {
        $fs.Dispose()
        throw
    }
    $Ctx.Archives[$LongPath] = $za
    return $za
}

function Test-UnsafeEntryName {
    # Zip slip protection: entries with absolute paths or ".." parts could write
    # outside the output folder.
    param([string]$Name)
    if ($Name -match '^[\\/]') { return $true }
    if ($Name -match '^[A-Za-z]:[\\/]') { return $true }
    foreach ($s in $Name.Split([char[]]@('/', '\'))) {
        if ($s -eq '..') { return $true }
    }
    return $false
}

function Test-IsClutter {
    param([string[]]$Segments, [bool]$IsDir)
    foreach ($s in $Segments) {
        if ($script:ClutterFolderNames -contains $s) { return $true }
    }
    if (-not $IsDir -and ($script:ClutterFileNames -contains $Segments[$Segments.Count - 1])) { return $true }
    return $false
}

function New-InventoryItem {
    param([string]$Kind, [string[]]$Segments, [string]$OrigRel)
    return [pscustomobject]@{
        Kind         = $Kind        # File, Dir, Container (an expanded zip) or Skipped
        Segments     = $Segments
        OrigRel      = $OrigRel
        Size         = [long]0
        LastWrite    = $null
        SourceType   = ''           # Fs or Zip
        SourcePath   = ''           # long path for Fs items
        ZipPath      = ''           # long path of the zip holding a Zip item
        EntryIndex   = -1
        Status       = ''           # set for skipped, rejected and failed items
        Note         = ''
        Node         = $null
        Hash         = ''
        IsZipTooDeep = $false
    }
}

function Add-DeclaredBytes {
    param($Ctx, [long]$Bytes)
    $Ctx.DeclaredBytes += $Bytes
    if ($Ctx.DeclaredBytes -gt $Ctx.MaxTotalBytes) {
        throw ("Stopped: the zip contents add up to {0:N2} GB uncompressed, which is over the {1} GB limit. Nothing was extracted. If this size is expected, run again with a higher -MaxTotalSizeGB." -f ($Ctx.DeclaredBytes / 1GB), $Ctx.MaxTotalSizeGB)
    }
}

function New-WorkFile {
    param($Ctx, [string]$Extension)
    if (-not $Ctx.WorkDir) {
        $name = '_work-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
        $Ctx.WorkDir = Join-PhysicalPath $Ctx.OutputLong @($name)
        [void][System.IO.Directory]::CreateDirectory($Ctx.WorkDir)
    }
    $Ctx.WorkCounter++
    return (Join-PhysicalPath $Ctx.WorkDir @(('n' + $Ctx.WorkCounter + $Extension)))
}

function Expand-NestedZip {
    # Reads a zip found inside the source. The inner zip is first copied to a
    # short temporary path in the output folder (never to its long original path).
    param($Ctx, $Item, [int]$Level, $Items)
    if ($Item.SourceType -eq 'Zip') {
        $tmp = New-WorkFile $Ctx '.zip'
        $entry = (Get-OpenArchive $Ctx $Item.ZipPath).Entries[$Item.EntryIndex]
        $in = $entry.Open()
        $out = [System.IO.File]::Open($tmp, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try { $r = Copy-StreamWithHash -InputStream $in -OutputStream $out -MaxBytes $Item.Size } finally { $out.Dispose(); $in.Dispose() }
        $Item.Hash = $r.Hash
        $zipLong = $tmp
    } else {
        $Item.Hash = Get-FileSha256 $Item.SourcePath
        $zipLong = $Item.SourcePath
    }
    $inner = New-Object System.Collections.Generic.List[object]
    Add-ZipInventory -Ctx $Ctx -ZipLongPath $zipLong -BaseSegments $Item.Segments -Level $Level -Items $inner
    $Item.Kind = 'Container'
    $Items.Add($Item)
    $Items.AddRange($inner)
}

function Add-FileOrNestedZip {
    param($Ctx, $Item, [int]$ContainerLevel, $Items)
    $name = $Item.Segments[$Item.Segments.Count - 1]
    if ($Ctx.ExpandNestedZips -and $name -match '\.zip$') {
        $level = $ContainerLevel + 1
        if ($level -le $script:MaxNestedZipLevel) {
            try {
                Expand-NestedZip -Ctx $Ctx -Item $Item -Level $level -Items $Items
                return
            } catch {
                if ($_.Exception.Message -like 'Stopped:*') { throw }
                $Item.Kind = 'File'
                $Item.Note = 'Could not open the zip inside, so it was kept as a zip file'
                Add-RunWarning $Ctx ("Could not open the zip inside '{0}': {1}" -f $Item.OrigRel, $_.Exception.Message)
            }
        } else {
            $Item.IsZipTooDeep = $true
        }
    }
    $Items.Add($Item)
}

function Add-ZipInventory {
    param($Ctx, [string]$ZipLongPath, [string[]]$BaseSegments, [int]$Level, $Items)
    $flags = Get-ZipEntryFlag -LongPath $ZipLongPath
    $archive = Get-OpenArchive $Ctx $ZipLongPath (Get-ZipNameEncoding $flags)
    $entries = $archive.Entries
    if ($null -eq $flags -or $flags.Count -ne $entries.Count) {
        $flags = $null
        Add-RunWarning $Ctx ("Could not read the index of '{0}' to check for encrypted entries. Encrypted entries will show as errors." -f ($(if ($BaseSegments.Count) { $BaseSegments -join '\' } else { 'the source zip' })))
    }

    # Check the declared size of everything in this zip before extracting anything.
    $declared = [long]0
    foreach ($e in $entries) { $declared += $e.Length }
    Add-DeclaredBytes $Ctx $declared

    $basePrefix = ''
    if ($BaseSegments.Count -gt 0) { $basePrefix = ($BaseSegments -join '\') + '\' }

    for ($i = 0; $i -lt $entries.Count; $i++) {
        Update-RunProgress ('Reading the zip ({0} of {1} items)' -f ($i + 1), $entries.Count) ($i + 1) $entries.Count
        $e = $entries[$i]
        $raw = $e.FullName

        if (Test-UnsafeEntryName $raw) {
            $it = New-InventoryItem 'Skipped' $null ($basePrefix + $raw.Replace('/', '\'))
            $it.Status = 'Rejected - unsafe path (zip slip)'
            $it.Note = 'Entry uses an absolute path or .. to point outside the output folder'
            $it.Size = $e.Length
            $Items.Add($it)
            continue
        }

        $segList = New-Object System.Collections.Generic.List[string]
        foreach ($s in $raw.Split([char[]]@('/', '\'))) { if ($s -ne '' -and $s -ne '.') { $segList.Add($s) } }
        if ($segList.Count -eq 0) { continue }
        $isDir = ($raw.EndsWith('/') -or $raw.EndsWith('\'))
        $segs = [string[]]($BaseSegments + $segList.ToArray())
        $origRel = $segs -join '\'

        if (-not $Ctx.IncludeSystemFiles -and (Test-IsClutter $segList.ToArray() $isDir)) {
            $it = New-InventoryItem 'Skipped' $segs $origRel
            $it.Status = 'Skipped - system file'
            $it.Note = 'Mac or Windows housekeeping file, not a document'
            $it.Size = $e.Length
            $Items.Add($it)
            continue
        }

        if ($null -ne $flags -and ((($flags[$i].Flags -band 1) -ne 0) -or $flags[$i].Method -eq 99)) {
            $it = New-InventoryItem 'Skipped' $segs $origRel
            $it.Status = 'Skipped - encrypted'
            $it.Note = 'Password protected entry. Extract it by hand with the password.'
            $it.Size = $e.Length
            $Items.Add($it)
            continue
        }

        $lastWrite = $null
        try { $lastWrite = $e.LastWriteTime.DateTime } catch { $lastWrite = $null }

        if ($isDir) {
            $it = New-InventoryItem 'Dir' $segs $origRel
            $it.LastWrite = $lastWrite
            $Items.Add($it)
            continue
        }

        $it = New-InventoryItem 'File' $segs $origRel
        $it.SourceType = 'Zip'
        $it.ZipPath = $ZipLongPath
        $it.EntryIndex = $i
        $it.Size = $e.Length
        $it.LastWrite = $lastWrite
        Add-FileOrNestedZip -Ctx $Ctx -Item $it -ContainerLevel $Level -Items $Items
    }
}

# ---------------------------------------------------------------------------
# Folder reading
# ---------------------------------------------------------------------------

function Get-LinkTypeSafe {
    param($Info)
    try {
        $t = [Microsoft.PowerShell.Commands.InternalSymbolicLinkLinkCodeMethods]::GetLinkType([psobject]$Info)
        if ($t) { return [string]$t }
        return ''
    } catch {
        return '?'
    }
}

function Add-FolderInventory {
    param($Ctx, [string]$RootLong, $Items)
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push([pscustomobject]@{ Info = (New-Object System.IO.DirectoryInfo -ArgumentList $RootLong); Segs = [string[]]@() })
    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()
        Update-RunProgress ('Reading the folder ({0} items so far)' -f $Items.Count)
        try {
            $children = @($cur.Info.EnumerateFileSystemInfos() | Sort-Object -Property Name)
        } catch {
            $rel = $cur.Segs -join '\'
            $it = New-InventoryItem 'Skipped' $cur.Segs $rel
            $it.Status = 'Error - could not read folder'
            $it.Note = $_.Exception.Message
            $Items.Add($it)
            continue
        }
        $subDirs = New-Object System.Collections.Generic.List[object]
        foreach ($c in $children) {
            $segs = [string[]]($cur.Segs + $c.Name)
            $origRel = $segs -join '\'
            $isDir = ($c -is [System.IO.DirectoryInfo])
            $attr = [int]$c.Attributes

            if (($attr -band 0x400) -ne 0) {
                # Reparse point. OneDrive files are reparse points too, so only
                # real links (junctions and symbolic links) are skipped.
                $lt = Get-LinkTypeSafe $c
                if ($lt -eq 'Junction' -or $lt -eq 'SymbolicLink' -or ($isDir -and $lt -eq '?')) {
                    $it = New-InventoryItem 'Skipped' $segs $origRel
                    $it.Status = 'Skipped - shortcut or link'
                    $it.Note = 'Links are not followed, so the tool cannot wander outside the source'
                    $Items.Add($it)
                    continue
                }
            }

            # Offline, RecallOnOpen or RecallOnDataAccess: reading this would make
            # OneDrive download it.
            if (($attr -band 0x441000) -ne 0) {
                $it = New-InventoryItem 'Skipped' $segs $origRel
                $it.Status = 'Skipped - online-only file'
                $it.Note = 'Not stored on this computer. Right-click it and choose "Always keep on this device", then run again.'
                if (-not $isDir) { $it.Size = $c.Length }
                $Items.Add($it)
                continue
            }

            if (-not $Ctx.IncludeSystemFiles -and (Test-IsClutter @($c.Name) $isDir)) {
                $it = New-InventoryItem 'Skipped' $segs $origRel
                $it.Status = 'Skipped - system file'
                $it.Note = 'Mac or Windows housekeeping file, not a document'
                if (-not $isDir) { $it.Size = $c.Length }
                $Items.Add($it)
                continue
            }

            if ($isDir) {
                $it = New-InventoryItem 'Dir' $segs $origRel
                $it.LastWrite = $c.LastWriteTimeUtc
                $Items.Add($it)
                $subDirs.Add([pscustomobject]@{ Info = $c; Segs = $segs })
                continue
            }

            $it = New-InventoryItem 'File' $segs $origRel
            $it.SourceType = 'Fs'
            $it.SourcePath = $c.FullName
            $it.Size = $c.Length
            $it.LastWrite = $c.LastWriteTimeUtc
            Add-FileOrNestedZip -Ctx $Ctx -Item $it -ContainerLevel 0 -Items $Items
        }
        for ($i = $subDirs.Count - 1; $i -ge 0; $i--) { $stack.Push($subDirs[$i]) }
    }
}

# ---------------------------------------------------------------------------
# The rename plan
# ---------------------------------------------------------------------------

function New-PlanNode {
    param($Parent, [string]$OrigName, [bool]$IsDir, [int]$Order)
    return [pscustomobject]@{
        OrigName   = $OrigName
        CleanBase  = ''
        Base       = ''
        Ext        = ''
        Suffix     = ''
        IsDir      = $IsDir
        IsRoot     = $false
        Parent     = $Parent
        Children   = (New-Object System.Collections.Generic.List[object])
        ChildIndex = (New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase))
        Rules      = (New-Object System.Collections.Generic.List[string])
        Applied    = @{}
        Order      = $Order
        Item       = $null
        DirItem    = $null
        IsContainer = $false
        Flagged    = $false
    }
}

function Add-NodeRule {
    param($Node, [string]$Label)
    if (-not $Node.Rules.Contains($Label)) { $Node.Rules.Add($Label) }
}

function Get-NodeName {
    param($Node)
    return ($Node.Base + $Node.Suffix + $Node.Ext)
}

function Get-NodeLength {
    # Full length of the node's path once it sits under the destination prefix.
    param($Ctx, $Node)
    $len = $Ctx.Prefix.Length
    $cur = $Node
    while (-not $cur.IsRoot) {
        $len += 1 + $cur.Base.Length + $cur.Suffix.Length + $cur.Ext.Length
        $cur = $cur.Parent
    }
    return $len
}

function Get-NodeDepth {
    param($Node)
    $d = 0
    $cur = $Node
    while (-not $cur.IsRoot) { $d++; $cur = $cur.Parent }
    return $d
}

function Get-NodeNames {
    param($Node)
    $names = New-Object System.Collections.Generic.List[string]
    $cur = $Node
    while (-not $cur.IsRoot) { $names.Insert(0, (Get-NodeName $cur)); $cur = $cur.Parent }
    return , $names.ToArray()
}

function Get-PathRules {
    param($Node)
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $cur = $Node
    while (-not $cur.IsRoot) {
        foreach ($r in $cur.Rules) { [void]$set.Add($r) }
        $cur = $cur.Parent
    }
    $ordered = foreach ($l in $script:RuleLabelOrder) { if ($set.Contains($l)) { $l } }
    return (@($ordered) -join ' + ')
}

function Build-PlanTree {
    param($Ctx)
    $root = New-PlanNode $null '' $true 0
    $root.IsRoot = $true
    $root.Base = $Ctx.PrefixLeaf
    $root.CleanBase = $Ctx.PrefixLeaf
    $order = 1
    foreach ($item in $Ctx.Items) {
        if ($item.Kind -eq 'Skipped') { continue }
        $parent = $root
        $segs = $item.Segments
        for ($i = 0; $i -lt $segs.Count; $i++) {
            $seg = $segs[$i]
            $isLast = ($i -eq $segs.Count - 1)
            if ($isLast -and $item.Kind -eq 'File') {
                $node = New-PlanNode $parent $seg $false $order
                $order++
                $node.Item = $item
                $parent.Children.Add($node)
                $item.Node = $node
                break
            }
            if ($parent.ChildIndex.ContainsKey($seg)) {
                $node = $parent.ChildIndex[$seg]
            } else {
                $node = New-PlanNode $parent $seg $true $order
                $order++
                $parent.Children.Add($node)
                $parent.ChildIndex[$seg] = $node
            }
            if ($isLast) {
                $item.Node = $node
                if ($item.Kind -eq 'Container') { $node.IsContainer = $true; $node.DirItem = $item }
                elseif ($null -eq $node.DirItem) { $node.DirItem = $item }
            }
            $parent = $node
        }
    }
    return $root
}

function Initialize-NodeName {
    # Rule 1 for every name, plus the 240 character cap on any single name.
    param($Node)
    $clean = Get-CleanName $Node.OrigName
    $name = $clean.Name
    foreach ($c in $clean.Changes) { Add-NodeRule $Node $c }
    if ($Node.IsContainer) {
        $name = ($name -replace '\.zip$', '').TrimEnd(' ', '.')
        if ($name -eq '') { $name = '_' }
        Add-NodeRule $Node 'Nested zip expanded'
    }
    if ($Node.IsDir) {
        $Node.Base = $name
        $Node.Ext = ''
    } else {
        $parts = Split-FileName $name
        $Node.Base = $parts.Stem
        $Node.Ext = $parts.Ext
    }
    if (($Node.Base.Length + $Node.Ext.Length) -gt $script:MaxNameLength) {
        if ($Node.IsDir) {
            $Node.Base = Get-TruncatedText $Node.Base $script:MaxNameLength $script:MinFolderLength
        } else {
            $prot = Get-ProtectedPrefix $Node.Base
            $restMax = $script:MaxNameLength - $Node.Ext.Length - $prot.Length
            $Node.Base = ($prot + (Get-TruncatedText ($Node.Base.Substring($prot.Length)) $restMax 1)).TrimEnd(' ', '.')
        }
        Add-NodeRule $Node 'Very long name shortened'
    }
    $Node.CleanBase = $Node.Base
}

function Get-OverLongLeaves {
    param($Ctx, $Leaves)
    $over = New-Object System.Collections.Generic.List[object]
    foreach ($l in $Leaves) {
        if ($Ctx.Hopeless.Contains($l.Order)) { continue }
        $len = Get-NodeLength $Ctx $l
        if ($len -gt $Ctx.Budget) { $over.Add([pscustomobject]@{ Node = $l; Len = $len }) }
    }
    return @($over | Sort-Object -Property @{ Expression = { $_.Len }; Descending = $true }, @{ Expression = { $_.Node.Order }; Descending = $false } | ForEach-Object { $_.Node })
}

function Get-SoftRuleResult {
    param($Ctx, $Node, [string]$Rule)
    $result = $null
    switch ($Rule) {
        'Abbreviations' {
            if ($Ctx.Abbreviations.Count -eq 0) { return $null }
            if ($Node.IsDir) {
                $result = Invoke-AbbreviationRule $Node.Base $Ctx.Abbreviations
            } else {
                $prot = Get-ProtectedPrefix $Node.Base
                $result = $prot + (Invoke-AbbreviationRule ($Node.Base.Substring($prot.Length)) $Ctx.Abbreviations)
            }
        }
        'Repeated parent name removed' {
            $parentNames = @($Node.Parent.Base, $Node.Parent.CleanBase)
            if ($Node.IsDir) {
                $result = Remove-ParentName $Node.Base $parentNames
            } else {
                $prot = Get-ProtectedPrefix $Node.Base
                $rest = Remove-ParentName ($Node.Base.Substring($prot.Length)) $parentNames
                if ($null -ne $rest) { $result = $prot + $rest }
            }
        }
        'Filler words removed' {
            if (-not $Node.IsDir) { return $null }
            $result = Remove-FillerWord $Node.Base
        }
    }
    if ($null -eq $result) { return $null }
    $result = (Get-CleanName $result).Name
    if ($result -eq '_' -or -not (Test-HasWordCharacter $result)) { return $null }
    if ($result.Length -ge $Node.Base.Length) { return $null }
    return $result
}

function Invoke-SoftRule {
    # Rules 2 to 4. Only names on a path that is too long are touched, the
    # biggest saving first, and the rule stops as soon as the path fits.
    param($Ctx, $Leaves, [string]$Rule)
    foreach ($leaf in (Get-OverLongLeaves $Ctx $Leaves)) {
        if ((Get-NodeLength $Ctx $leaf) -le $Ctx.Budget) { continue }
        $cands = New-Object System.Collections.Generic.List[object]
        $cur = $leaf
        while (-not $cur.IsRoot) {
            if (-not $cur.Applied.ContainsKey($Rule)) {
                $new = Get-SoftRuleResult $Ctx $cur $Rule
                if ($null -ne $new) {
                    $cands.Add([pscustomobject]@{ Node = $cur; NewBase = $new; Saving = $cur.Base.Length - $new.Length; Depth = (Get-NodeDepth $cur) })
                }
            }
            $cur = $cur.Parent
        }
        $sorted = @($cands | Sort-Object -Property @{ Expression = { $_.Saving }; Descending = $true }, @{ Expression = { $_.Depth }; Descending = $false })
        foreach ($c in $sorted) {
            if ((Get-NodeLength $Ctx $leaf) -le $Ctx.Budget) { break }
            $c.Node.Base = $c.NewBase
            $c.Node.Applied[$Rule] = $true
            Add-NodeRule $c.Node $Rule
        }
    }
}

function Invoke-FolderTruncation {
    # Rule 5. Cuts the longest folder names on a too-long path first, levelling
    # them down towards the next longest, never below 12 characters.
    param($Ctx, $Leaves)
    foreach ($leaf in (Get-OverLongLeaves $Ctx $Leaves)) {
        $guard = 0
        while ($guard -lt 1000) {
            $guard++
            $excess = (Get-NodeLength $Ctx $leaf) - $Ctx.Budget
            if ($excess -le 0) { break }
            $cands = New-Object System.Collections.Generic.List[object]
            if ($leaf.IsDir) { $cur = $leaf } else { $cur = $leaf.Parent }
            while (-not $cur.IsRoot) {
                if ($cur.Base.Length -gt $script:MinFolderLength) { $cands.Add($cur) }
                $cur = $cur.Parent
            }
            if ($cands.Count -eq 0) { break }
            $sorted = @($cands | Sort-Object -Property @{ Expression = { $_.Base.Length }; Descending = $true }, @{ Expression = { Get-NodeDepth $_ }; Descending = $false })
            $longest = $sorted[0].Base.Length
            $tied = @($sorted | Where-Object { $_.Base.Length -eq $longest })
            if ($tied.Count -eq 1) {
                $second = $script:MinFolderLength
                if ($sorted.Count -gt 1) { $second = [Math]::Max($second, $sorted[1].Base.Length) }
                $target = [Math]::Max($second, $longest - $excess)
            } else {
                $share = [int][Math]::Ceiling($excess / $tied.Count)
                $target = [Math]::Max($script:MinFolderLength, $longest - $share)
            }
            foreach ($t in $tied) {
                $nb = Get-TruncatedText $t.Base $target $script:MinFolderLength
                if ($nb.Length -ge $t.Base.Length) { $nb = $t.Base.Substring(0, $t.Base.Length - 1).TrimEnd(' ', '.') }
                if ($nb.Length -eq 0) { $nb = '_' }
                $t.Base = $nb
                Add-NodeRule $t 'Folder name shortened'
                if ((Get-NodeLength $Ctx $leaf) -le $Ctx.Budget) { break }
            }
        }
    }
}

function Invoke-FileTruncation {
    # Rule 6. Cuts file names at a word boundary, keeping the extension and any
    # leading date or document number.
    param($Ctx, $Leaves)
    foreach ($leaf in (Get-OverLongLeaves $Ctx $Leaves)) {
        if ($leaf.IsDir) { continue }
        $excess = (Get-NodeLength $Ctx $leaf) - $Ctx.Budget
        if ($excess -le 0) { continue }
        $prot = Get-ProtectedPrefix $leaf.Base
        $rest = $leaf.Base.Substring($prot.Length)
        $minStem = [Math]::Max($script:MinFileStemLength, $prot.Length)
        if ($leaf.Base.Length -le $minStem -or $rest.Length -eq 0) { continue }
        $targetStem = [Math]::Max($minStem, $leaf.Base.Length - $excess)
        $restMax = $targetStem - $prot.Length
        $restMin = [Math]::Max(1, $minStem - $prot.Length)
        $newRest = Get-TruncatedText $rest $restMax $restMin
        $newBase = ($prot + $newRest).TrimEnd(' ', '.', '-', '_', ',')
        if ($newBase.Length -lt $prot.TrimEnd().Length) { $newBase = $prot.TrimEnd() }
        if ($newBase.Length -gt 0 -and $newBase.Length -lt $leaf.Base.Length) {
            $leaf.Base = $newBase
            Add-NodeRule $leaf 'File name shortened'
        }
    }
}

function Resolve-NameCollision {
    # Windows and SharePoint treat names that differ only in case as the same.
    # Duplicates get " (2)", " (3)" and so on. A name the client sent unchanged
    # keeps priority over a name the tool changed. Once a node is numbered it
    # keeps its number, even if it is shortened again later to make room for it.
    param($Root)
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        foreach ($c in $dir.Children) { if ($c.IsDir) { $stack.Push($c) } }
        if ($dir.Children.Count -lt 2) { continue }

        $ordered = @($dir.Children | Sort-Object -Property @{ Expression = {
                    if ((@($_.Rules | Where-Object { $_ -ne $script:DuplicateLabel })).Count -eq 0) { 0 } else { 1 } } }, @{ Expression = { $_.Order } })
        $taken = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in $ordered) {
            if ($taken.Add($m.Base + $m.Suffix + $m.Ext)) { continue }
            $n = 2
            while ($taken.Contains($m.Base + ' (' + $n + ')' + $m.Ext)) { $n++ }
            $m.Suffix = ' (' + $n + ')'
            [void]$taken.Add($m.Base + $m.Suffix + $m.Ext)
            Add-NodeRule $m $script:DuplicateLabel
        }
    }
}

function Invoke-RenameRule {
    # Runs rules 1 to 6 and the duplicate check over a freshly built tree.
    param($Ctx)
    $root = Build-PlanTree $Ctx
    $all = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[object]
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $n = $stack.Pop()
        if (-not $n.IsRoot) { $all.Add($n) }
        foreach ($c in $n.Children) { $stack.Push($c) }
    }
    $leaves = @($all | Where-Object { (-not $_.IsDir) -or $_.Children.Count -eq 0 } | Sort-Object -Property Order)

    foreach ($n in $all) { Initialize-NodeName $n }
    Resolve-NameCollision $root

    foreach ($rule in 'Abbreviations', 'Repeated parent name removed', 'Filler words removed') {
        Invoke-SoftRule $Ctx $leaves $rule
        Resolve-NameCollision $root
    }
    for ($pass = 0; $pass -lt 6; $pass++) {
        if (@(Get-OverLongLeaves $Ctx $leaves).Count -eq 0) { break }
        Invoke-FolderTruncation $Ctx $leaves
        Resolve-NameCollision $root
        Invoke-FileTruncation $Ctx $leaves
        Resolve-NameCollision $root
    }
    $Ctx.Root = $root
    $Ctx.AllNodes = $all
    $Ctx.Leaves = $leaves
}

function New-RenamePlan {
    # Pass 1 finds the paths that cannot fit however much they are shortened.
    # Pass 2 starts again and shortens only for paths that can be fixed, so a
    # hopeless path never causes folder names shared with other files to be cut
    # for nothing. Hopeless paths are flagged for manual attention.
    param($Ctx)
    $Ctx.Hopeless = New-Object 'System.Collections.Generic.HashSet[int]'
    Invoke-RenameRule $Ctx
    $hopeless = @($Ctx.Leaves | Where-Object { (Get-NodeLength $Ctx $_) -gt $Ctx.Budget })
    if ($hopeless.Count -gt 0) {
        foreach ($l in $hopeless) { [void]$Ctx.Hopeless.Add($l.Order) }
        Invoke-RenameRule $Ctx
    }
    foreach ($l in $Ctx.Leaves) { $l.Flagged = ((Get-NodeLength $Ctx $l) -gt $Ctx.Budget) }
}

# ---------------------------------------------------------------------------
# Report rows
# ---------------------------------------------------------------------------

function Get-StatusRank {
    param([string]$Status)
    if ($Status -eq 'Needs manual attention') { return 0 }
    if ($Status -like 'Error*') { return 1 }
    if ($Status -like 'Rejected*') { return 2 }
    if ($Status -like 'Skipped*') { return 3 }
    if ($Status -eq 'Renamed') { return 4 }
    if ($Status -like 'Expanded*') { return 5 }
    return 6
}

function Get-ReportRows {
    param($Ctx)
    $rows = New-Object System.Collections.Generic.List[object]
    $seenNodes = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($item in $Ctx.Items) {
        $origFull = $Ctx.Prefix + '\' + $item.OrigRel
        if ($item.Kind -eq 'Skipped') {
            $rows.Add([pscustomobject]@{
                    OriginalPath = $origFull; NewPath = ''; OriginalLength = $origFull.Length; NewLength = ''
                    RulesApplied = $item.Note; Status = $item.Status
                })
            continue
        }
        $node = $item.Node
        if ($null -eq $node) { continue }
        if ($item.Kind -eq 'Dir' -and ($node.Children.Count -gt 0 -or $node.IsContainer)) { continue }
        if (-not $seenNodes.Add($node.Order)) { continue }

        $newRel = (Get-NodeNames $node) -join '\'
        $newFull = $Ctx.Prefix + '\' + $newRel
        $rules = Get-PathRules $node
        if ($item.Kind -eq 'Dir') { $rules = (@('Empty folder') + @($rules | Where-Object { $_ })) -join ' + ' }
        if ($item.IsZipTooDeep) { $rules = (@($rules | Where-Object { $_ }) + @('Zip nested more than 3 levels deep, kept as a zip')) -join ' + ' }

        if ($item.Status) { $status = $item.Status }
        elseif ($node.Flagged) { $status = 'Needs manual attention' }
        elseif ($item.Kind -eq 'Container') { $status = 'Expanded (nested zip)' }
        elseif ([string]::Equals($newRel, $item.OrigRel, [System.StringComparison]::Ordinal)) { $status = 'OK' }
        else { $status = 'Renamed' }

        $rows.Add([pscustomobject]@{
                OriginalPath = $origFull; NewPath = $newFull; OriginalLength = $origFull.Length; NewLength = $newFull.Length
                RulesApplied = $rules; Status = $status
            })
    }
    return @($rows | Sort-Object -Property @{ Expression = { Get-StatusRank $_.Status } }, @{ Expression = { $_.OriginalPath } })
}

# ---------------------------------------------------------------------------
# Apply: copy or extract with the new names
# ---------------------------------------------------------------------------

function Copy-ItemToTarget {
    param($Ctx, $Item, [string]$TargetLong)
    $in = $null
    $out = $null
    try {
        if ($Item.SourceType -eq 'Fs') {
            $in = [System.IO.File]::Open($Item.SourcePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            $max = -1
        } else {
            $entry = (Get-OpenArchive $Ctx $Item.ZipPath).Entries[$Item.EntryIndex]
            $in = $entry.Open()
            $max = $Item.Size
        }
        # CreateNew: never overwrite anything that is already there.
        $out = [System.IO.File]::Open($TargetLong, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $r = Copy-StreamWithHash -InputStream $in -OutputStream $out -MaxBytes $max
    } finally {
        if ($null -ne $out) { $out.Dispose() }
        if ($null -ne $in) { $in.Dispose() }
    }
    if ($Item.SourceType -eq 'Zip' -and $r.Bytes -ne $Item.Size) {
        throw ('The entry holds {0} bytes but the zip says {1}. The zip may be damaged.' -f $r.Bytes, $Item.Size)
    }
    if ($null -ne $Item.LastWrite) { [System.IO.File]::SetLastWriteTime($TargetLong, [datetime]$Item.LastWrite) }
    $check = Get-FileSha256 $TargetLong
    if ($check -ne $r.Hash) {
        throw ('HASH MISMATCH: {0} was written with different contents from the source (source {1}, output {2}). The run was stopped.' -f $Item.OrigRel, $r.Hash, $check)
    }
    return $r
}

function Get-AttentionName {
    # Files that still do not fit are saved flat in the "Needs attention" folder,
    # with names short enough that File Explorer, Word and Excel can open them
    # there. The report shows where each one was meant to go.
    param($Ctx, $Leaves)
    $map = @{}
    $rootFull = ConvertFrom-LongPath $Ctx.AttentionLong
    $maxName = [Math]::Max(40, 250 - $rootFull.Length - 1)
    $used = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($leaf in $Leaves) {
        if (-not $leaf.Flagged -or $leaf.IsDir) { continue }
        $stem = $leaf.Base
        $ext = $leaf.Ext
        if (($stem + $ext).Length -gt $maxName - 6) {
            $prot = Get-ProtectedPrefix $stem
            $restMax = [Math]::Max(1, $maxName - 6 - $ext.Length - $prot.Length)
            $stem = ($prot + (Get-TruncatedText ($stem.Substring($prot.Length)) $restMax 1)).TrimEnd(' ', '.')
        }
        $name = $stem + $ext
        $n = 2
        while (-not $used.Add($name)) {
            $name = $stem + ' (' + $n + ')' + $ext
            $n++
        }
        $map[$leaf.Order] = $name
    }
    return $map
}

function Invoke-ApplyPlan {
    param($Ctx)
    $log = $Ctx.LogRows

    if ($Ctx.SourceIsZip) {
        $fi = New-Object System.IO.FileInfo -ArgumentList $Ctx.SourceLong
        $log.Add([pscustomobject]@{
                OriginalRelativePath = $fi.Name; NewRelativePath = ''; SizeBytes = $fi.Length
                LastModified = (Format-LocalTime $fi.LastWriteTimeUtc); SHA256 = (Get-FileSha256 $Ctx.SourceLong)
                Result = 'Source zip (opened read only, not changed)'
            })
    }

    [void][System.IO.Directory]::CreateDirectory($Ctx.FilesLong)
    $ordered = @($Ctx.Leaves | Sort-Object -Property @{ Expression = { (Get-NodeNames $_) -join '\' } })
    $attentionNames = Get-AttentionName $Ctx $ordered
    $done = 0
    foreach ($leaf in $ordered) {
        if ($leaf.Flagged) {
            # Empty folders that are too long are listed in the report only.
            if ($leaf.IsDir) { continue }
            $rootLong = $Ctx.AttentionLong
            $names = [string[]]@($attentionNames[$leaf.Order])
        } else {
            $rootLong = $Ctx.FilesLong
            $names = Get-NodeNames $leaf
        }
        $target = Join-PhysicalPath $rootLong $names
        if (-not (Test-PathUnder $target $rootLong) -or $target.Length -le $rootLong.Length) {
            throw ('Refusing to write outside the output folder: {0}' -f $target)
        }
        if ($leaf.IsDir) {
            [void][System.IO.Directory]::CreateDirectory($target)
            continue
        }
        $item = $leaf.Item
        $newRel = $names -join '\'
        Update-RunProgress ('Copying files ({0} of {1})' -f ($done + 1), $ordered.Count) $done $ordered.Count
        [void][System.IO.Directory]::CreateDirectory((Get-ParentPhysicalPath $target))
        try {
            $r = Copy-ItemToTarget $Ctx $item $target
            $item.Hash = $r.Hash
            if ($leaf.Flagged) { $result = 'Copied to Needs attention folder' } else { $result = 'Copied' }
            $log.Add([pscustomobject]@{
                    OriginalRelativePath = $item.OrigRel; NewRelativePath = $newRel; SizeBytes = $r.Bytes
                    LastModified = (Format-LocalTime $item.LastWrite); SHA256 = $r.Hash; Result = $result
                })
        } catch {
            $msg = $_.Exception.Message
            # A mismatched file is kept as evidence. Anything else half written is removed.
            if ($msg -like 'HASH MISMATCH*') { throw }
            try { if ([System.IO.File]::Exists($target)) { [System.IO.File]::Delete($target) } } catch { }
            if ($msg -like 'Stopped:*') { throw }
            $item.Status = 'Error - ' + $msg
            $log.Add([pscustomobject]@{
                    OriginalRelativePath = $item.OrigRel; NewRelativePath = ''; SizeBytes = $item.Size
                    LastModified = (Format-LocalTime $item.LastWrite); SHA256 = ''; Result = ('Error - ' + $msg)
                })
            Add-RunWarning $Ctx ('Could not copy {0}: {1}' -f $item.OrigRel, $msg)
        }
        $done++
        if (-not $script:Quiet -and ($done % 200) -eq 0) { Write-Info ('  ...{0} of {1} files written' -f $done, $ordered.Count) }
    }

    foreach ($item in $Ctx.Items) {
        if ($item.Kind -eq 'Container') {
            $log.Add([pscustomobject]@{
                    OriginalRelativePath = $item.OrigRel; NewRelativePath = ((Get-NodeNames $item.Node) -join '\'); SizeBytes = $item.Size
                    LastModified = (Format-LocalTime $item.LastWrite); SHA256 = $item.Hash; Result = 'Zip inside the source, expanded into this folder'
                })
        } elseif ($item.Kind -eq 'Skipped') {
            $log.Add([pscustomobject]@{
                    OriginalRelativePath = $item.OrigRel; NewRelativePath = ''; SizeBytes = $item.Size
                    LastModified = ''; SHA256 = ''; Result = $item.Status
                })
        }
    }

    # Folder dates, deepest first so writing a child does not disturb its parent.
    $dirs = @($Ctx.AllNodes | Where-Object { $_.IsDir -and $null -ne $_.DirItem -and $null -ne $_.DirItem.LastWrite } |
            Sort-Object -Property @{ Expression = { Get-NodeDepth $_ }; Descending = $true })
    foreach ($d in $dirs) {
        $names = Get-NodeNames $d
        foreach ($rootLong in @($Ctx.FilesLong, $Ctx.AttentionLong)) {
            $p = Join-PhysicalPath $rootLong $names
            if ([System.IO.Directory]::Exists($p)) {
                try { [System.IO.Directory]::SetLastWriteTime($p, [datetime]$d.DirItem.LastWrite) } catch { }
            }
        }
    }
}

function Get-ZipSafeDate {
    param($Date)
    if ($null -eq $Date) { $d = Get-Date } else { $d = [datetime]$Date }
    if ($d.Kind -eq [System.DateTimeKind]::Utc) { $d = $d.ToLocalTime() }
    if ($d.Year -lt 1980) { $d = New-Object DateTime -ArgumentList 1980, 1, 1, 0, 0, 0 }
    if ($d.Year -gt 2107) { $d = New-Object DateTime -ArgumentList 2107, 12, 31, 0, 0, 0 }
    return (New-Object System.DateTimeOffset -ArgumentList ([datetime]::SpecifyKind($d, [System.DateTimeKind]::Unspecified)), ([System.TimeZoneInfo]::Local.GetUtcOffset($d)))
}

function New-OutputZip {
    param($Ctx)
    $fs = [System.IO.File]::Open($Ctx.ZipLong, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $za = New-Object System.IO.Compression.ZipArchive -ArgumentList @($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        $ordered = @($Ctx.Leaves | Where-Object { -not $_.Flagged } | Sort-Object -Property @{ Expression = { (Get-NodeNames $_) -join '\' } })
        foreach ($leaf in $ordered) {
            $names = Get-NodeNames $leaf
            if ($leaf.IsDir) {
                $e = $za.CreateEntry((($names -join '/') + '/'))
                if ($null -ne $leaf.DirItem) { $e.LastWriteTime = Get-ZipSafeDate $leaf.DirItem.LastWrite }
                continue
            }
            if ($leaf.Item.Status) { continue }
            $src = Join-PhysicalPath $Ctx.FilesLong $names
            $e = $za.CreateEntry(($names -join '/'), [System.IO.Compression.CompressionLevel]::Optimal)
            $e.LastWriteTime = Get-ZipSafeDate $leaf.Item.LastWrite
            $in = [System.IO.File]::Open($src, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            $es = $e.Open()
            try { $in.CopyTo($es) } finally { $es.Dispose(); $in.Dispose() }
        }
    } finally {
        $za.Dispose()
    }
    $fi = New-Object System.IO.FileInfo -ArgumentList $Ctx.ZipLong
    $Ctx.LogRows.Add([pscustomobject]@{
            OriginalRelativePath = ''; NewRelativePath = $fi.Name; SizeBytes = $fi.Length
            LastModified = (Format-LocalTime $fi.LastWriteTimeUtc); SHA256 = (Get-FileSha256 $Ctx.ZipLong)
            Result = 'New zip of the renamed files'
        })
}

# ---------------------------------------------------------------------------
# Safety checks
# ---------------------------------------------------------------------------

function Get-SyncRoot {
    # Folders that OneDrive keeps in sync. The environment variables only cover
    # personal OneDrive, so the local list of synced SharePoint libraries in the
    # registry is read too. Read only and offline.
    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($name in 'OneDrive', 'OneDriveCommercial', 'OneDriveConsumer') {
        $v = [System.Environment]::GetEnvironmentVariable($name)
        if ($v) { $roots.Add($v) }
    }
    if ($script:IsWindowsOS) {
        foreach ($key in 'HKCU:\Software\SyncEngines\Providers\OneDrive', 'HKCU:\Software\Microsoft\OneDrive\Accounts') {
            try {
                foreach ($k in (Get-ChildItem -LiteralPath $key -ErrorAction Stop)) {
                    $p = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue
                    if ($null -eq $p) { continue }
                    $propNames = @($p.PSObject.Properties | ForEach-Object { $_.Name })
                    foreach ($prop in 'MountPoint', 'UserFolder') {
                        if (($propNames -contains $prop) -and $p.$prop) { $roots.Add([string]$p.$prop) }
                    }
                }
            } catch { }
        }
    }
    return , $roots.ToArray()
}

function Get-SyncedFolderMatch {
    <#
    .SYNOPSIS
        Returns the OneDrive or SharePoint synced folder that contains the
        output folder, or an empty string if it is not inside one. The
        destination folder counts as synced too.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [string]$DestinationPrefix = ''
    )
    $outFull = Resolve-FullPath $OutputFolder
    $prefix = $DestinationPrefix.Trim().Trim('"').Trim().Replace('/', '\')
    while ($prefix.Length -gt 3 -and $prefix.EndsWith('\')) { $prefix = $prefix.Substring(0, $prefix.Length - 1) }
    foreach ($r in @(Get-SyncRoot) + @($prefix)) {
        if ($r -and (Test-PathUnder $outFull $r)) { return [string]$r }
    }
    return ''
}

function Confirm-SyncedOutput {
    param([string]$OutputFull, [string]$SyncRoot)
    Write-Warning ("The output folder {0} is inside a OneDrive or SharePoint synced folder ({1}). Anything written there starts uploading straight away, including files that still need attention." -f $OutputFull, $SyncRoot)
    $answer = Read-Host 'Type YES to continue anyway, or press Enter to stop'
    if ($answer -ne 'YES') {
        throw 'Stopped. Choose an output folder outside OneDrive and SharePoint, for example C:\CL\Out.'
    }
}

function Test-LongPathSupport {
    # Creates, reads and deletes a path over 300 characters inside the output
    # folder, so a machine that blocks long paths fails here and not halfway.
    param([string]$OutputLong)
    $base = Join-PhysicalPath $OutputLong @(('_lpcheck-' + [guid]::NewGuid().ToString('N').Substring(0, 8)))
    $segs = @(1..6 | ForEach-Object { 'x' * 50 })
    $dir = Join-PhysicalPath $base $segs
    $file = Join-PhysicalPath $dir @('check.txt')
    try {
        [void][System.IO.Directory]::CreateDirectory($dir)
        [System.IO.File]::WriteAllText($file, 'ok')
        $ok = ([System.IO.File]::ReadAllText($file) -eq 'ok')
        return $ok
    } catch {
        return $false
    } finally {
        try { if ([System.IO.Directory]::Exists($base)) { [System.IO.Directory]::Delete($base, $true) } } catch { }
    }
}

# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

function Invoke-PathShortener {
    <#
    .SYNOPSIS
        Finds paths that will be too long once files sit in a synced SharePoint
        folder, and copies or extracts them with shorter names.
    .DESCRIPTION
        Dry run by default: writes a CSV report and changes nothing. Use -Apply
        to copy (folder source) or extract (zip source) into the output folder
        with the new names. The source is never modified.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$DestinationPrefix,
        [string]$OutputFolder = 'C:\CL\Out',
        [ValidateRange(30, 32000)][int]$MaxPathLength = 218,
        [ValidateRange(0, 1000)][int]$SafetyMargin = 10,
        [string]$AbbreviationsCsv = '',
        [switch]$Apply,
        [switch]$ExpandNestedZips,
        [switch]$CreateZip,
        [ValidateRange(0.0, 100000.0)][double]$MaxTotalSizeGB = 20,
        [switch]$AllowSyncedOutput,
        [switch]$IncludeSystemFiles,
        [switch]$Quiet,
        [hashtable]$Progress = $null
    )

    $script:Quiet = [bool]$Quiet
    $script:Progress = $Progress
    Update-RunProgress 'Checking the output folder...'
    $stamp = (Get-Date).ToString('yyyy-MM-dd HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)

    # Destination prefix: the final synced SharePoint folder.
    $prefix = $DestinationPrefix.Trim().Trim('"').Trim().Replace('/', '\')
    if ($prefix -notmatch '^[A-Za-z]:\\' -and $prefix -notmatch '^\\\\[^\\]+\\[^\\]+') {
        throw 'The destination folder must be a full path, for example C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd'
    }
    while ($prefix.Length -gt 3 -and $prefix.EndsWith('\')) { $prefix = $prefix.Substring(0, $prefix.Length - 1) }
    $budget = $MaxPathLength - $SafetyMargin
    if ($prefix.Length + 1 + $script:MinFolderLength -gt $budget) {
        throw ("The destination folder path is already {0} characters long, which leaves no room under the {1} character limit (with a {2} character safety margin). Pick a shorter destination folder." -f $prefix.Length, $MaxPathLength, $SafetyMargin)
    }
    $prefixParts = @($prefix.Split('\') | Where-Object { $_ -ne '' })
    $prefixLeaf = ''
    if ($prefixParts.Count -gt 1) { $prefixLeaf = $prefixParts[$prefixParts.Count - 1] }

    # Source.
    $srcFull = Resolve-FullPath $Source
    $srcLong = ConvertTo-LongPath $srcFull
    $srcIsDir = [System.IO.Directory]::Exists($srcLong)
    $srcIsFile = [System.IO.File]::Exists($srcLong)
    if (-not $srcIsDir -and -not $srcIsFile) { throw "Cannot find the source: $srcFull" }
    if ($srcIsFile -and $srcFull -notmatch '\.zip$') { throw "The source must be a folder or a .zip file: $srcFull" }

    # Output folder.
    $outFull = Resolve-FullPath $OutputFolder
    if ($srcIsDir -and ((Test-PathUnder $outFull $srcFull) -or (Test-PathUnder $srcFull $outFull))) {
        throw 'The output folder and the source folder overlap. Choose an output folder outside the source, for example C:\CL\Out.'
    }
    $outLong = ConvertTo-LongPath $outFull

    $syncHit = Get-SyncedFolderMatch -OutputFolder $outFull -DestinationPrefix $prefix
    $outputIsSynced = [bool]$syncHit
    if ($outputIsSynced -and -not $AllowSyncedOutput) { Confirm-SyncedOutput $outFull $syncHit }

    [void][System.IO.Directory]::CreateDirectory($outLong)
    if (-not (Test-LongPathSupport $outLong)) {
        throw ('This computer would not let the tool create a path longer than 260 characters in {0}. Nothing was changed. Ask IT to check that .NET Framework 4.6.2 or later is installed, or try an output folder on the local C: drive.' -f $outFull)
    }

    $abbrevs = @()
    if ($AbbreviationsCsv) {
        $abbrFull = Resolve-FullPath $AbbreviationsCsv
        $abbrLong = ConvertTo-LongPath $abbrFull
        if (-not [System.IO.File]::Exists($abbrLong)) { throw "Cannot find the abbreviations file: $abbrFull" }
        $abbrevs = @(Import-Abbreviation $abbrLong)
    }

    if ($srcIsFile) { $sourceName = [System.IO.Path]::GetFileNameWithoutExtension($srcFull) } else { $sourceName = [System.IO.Path]::GetFileName($srcFull) }
    $jobBase = Get-TruncatedText (Get-CleanName $sourceName).Name 40 8
    if (-not $jobBase) { $jobBase = 'Output' }

    $ctx = @{
        Prefix             = $prefix
        PrefixLeaf         = $prefixLeaf
        Budget             = $budget
        SourceFull         = $srcFull
        SourceLong         = $srcLong
        SourceIsZip        = $srcIsFile
        OutputFull         = $outFull
        OutputLong         = $outLong
        ExpandNestedZips   = [bool]$ExpandNestedZips
        IncludeSystemFiles = [bool]$IncludeSystemFiles
        MaxTotalSizeGB     = $MaxTotalSizeGB
        MaxTotalBytes      = [long]($MaxTotalSizeGB * 1GB)
        DeclaredBytes      = [long]0
        Abbreviations      = $abbrevs
        Items              = (New-Object System.Collections.Generic.List[object])
        Archives           = @{}
        Warnings           = (New-Object System.Collections.Generic.List[string])
        LogRows            = (New-Object System.Collections.Generic.List[object])
        WorkDir            = ''
        WorkCounter        = 0
        Root               = $null
        Hopeless           = (New-Object 'System.Collections.Generic.HashSet[int]')
        AllNodes           = $null
        Leaves             = @()
        FilesLong          = ''
        AttentionLong      = ''
        ZipLong            = ''
    }

    if ($Apply) { $modeText = 'APPLY' } else { $modeText = 'DRY RUN (nothing is copied or renamed)' }
    Write-Info ''
    Write-Info ('Long path shortener - ' + $modeText) 'Cyan'
    Write-Info ('Reading ' + $srcFull)

    $reportFull = ''
    $logFull = ''
    $filesFull = ''
    $attentionFull = ''
    $zipFull = ''
    $rows = @()
    try {
        if ($srcIsFile) {
            Add-ZipInventory -Ctx $ctx -ZipLongPath $srcLong -BaseSegments ([string[]]@()) -Level 0 -Items $ctx.Items
        } else {
            Add-FolderInventory -Ctx $ctx -RootLong $srcLong -Items $ctx.Items
        }

        Write-Info 'Working out new names...'
        Update-RunProgress 'Working out new names...'
        New-RenamePlan $ctx

        if ($Apply) {
            $n = 1
            while ($true) {
                if ($n -eq 1) { $job = $jobBase } else { $job = '{0} ({1})' -f $jobBase, $n }
                $f = Join-PhysicalPath $outLong @($job)
                $a = Join-PhysicalPath $outLong @(($job + ' - Needs attention'))
                $z = Join-PhysicalPath $outLong @(($job + '.zip'))
                if (-not (Test-NonEmptyDirectory $f) -and -not (Test-NonEmptyDirectory $a) -and -not ($CreateZip -and [System.IO.File]::Exists($z))) { break }
                $n++
            }
            $ctx.FilesLong = $f
            $ctx.AttentionLong = $a
            $filesFull = ConvertFrom-LongPath $f
            Write-Info ('Writing renamed files to ' + $filesFull)
            try {
                Invoke-ApplyPlan $ctx
                if ([System.IO.Directory]::Exists($a)) { $attentionFull = ConvertFrom-LongPath $a }
                if ($CreateZip) {
                    $ctx.ZipLong = $z
                    Write-Info 'Creating the new zip...'
                    Update-RunProgress 'Creating the new zip...'
                    New-OutputZip $ctx
                    $zipFull = ConvertFrom-LongPath $z
                }
            } finally {
                $logLong = Join-PhysicalPath $outLong @(('{0} - Log {1}.csv' -f $job, $stamp))
                Export-CsvFile $logLong @('OriginalRelativePath', 'NewRelativePath', 'SizeBytes', 'LastModified', 'SHA256', 'Result') $ctx.LogRows
                $logFull = ConvertFrom-LongPath $logLong
            }
            $reportLong = Join-PhysicalPath $outLong @(('{0} - Report {1}.csv' -f $job, $stamp))
        } else {
            $reportLong = Join-PhysicalPath $outLong @(('{0} - Dry run report {1}.csv' -f $jobBase, $stamp))
        }

        Update-RunProgress 'Writing the report...'
        $rows = Get-ReportRows $ctx
        Export-CsvFile $reportLong @('OriginalPath', 'NewPath', 'OriginalLength', 'NewLength', 'RulesApplied', 'Status') $rows
        $reportFull = ConvertFrom-LongPath $reportLong
    } finally {
        foreach ($za in $ctx.Archives.Values) { try { $za.Dispose() } catch { } }
        if ($ctx.WorkDir -and [System.IO.Directory]::Exists($ctx.WorkDir)) {
            try { [System.IO.Directory]::Delete($ctx.WorkDir, $true) } catch { }
        }
    }

    $processed = @($ctx.Items | Where-Object { $_.Kind -eq 'File' }).Count
    $renamed = @($rows | Where-Object { $_.Status -eq 'Renamed' }).Count
    $flagged = @($rows | Where-Object { $_.Status -eq 'Needs manual attention' }).Count
    $skipped = @($rows | Where-Object { $_.Status -like 'Skipped*' }).Count
    $rejected = @($rows | Where-Object { $_.Status -like 'Rejected*' }).Count
    $errors = @($rows | Where-Object { $_.Status -like 'Error*' }).Count
    $expanded = @($rows | Where-Object { $_.Status -like 'Expanded*' }).Count

    Write-Info ''
    Write-Info '------------------------------------------------------------' 'Cyan'
    if ($Apply) { Write-Info ' Summary' 'Cyan' } else { Write-Info ' Summary (dry run - nothing was copied or renamed)' 'Cyan' }
    Write-Info '------------------------------------------------------------' 'Cyan'
    Write-Info (' Destination folder:   {0}' -f $prefix)
    Write-Info (' Path limit:           {0} characters, less a {1} character safety margin = {2}' -f $MaxPathLength, $SafetyMargin, $budget)
    Write-Info (' Files processed:      {0}' -f $processed)
    Write-Info (' Files renamed:        {0}' -f $renamed)
    if ($flagged -gt 0) {
        Write-Info (' Needing attention:    {0}  (still too long, see Status "Needs manual attention" in the report)' -f $flagged) 'Yellow'
    } else {
        Write-Info (' Needing attention:    0')
    }
    if ($expanded -gt 0) { Write-Info (' Zips inside expanded: {0}' -f $expanded) }
    if ($skipped -gt 0) { Write-Info (' Skipped:              {0}  (see the report for why)' -f $skipped) }
    if ($rejected -gt 0) { Write-Info (' Rejected as unsafe:   {0}  (entries that tried to write outside the output folder)' -f $rejected) 'Yellow' }
    if ($errors -gt 0) { Write-Info (' Errors:               {0}' -f $errors) 'Red' }
    Write-Info (' Report saved to:      {0}' -f $reportFull)
    if ($Apply) {
        Write-Info (' Evidence log saved to: {0}' -f $logFull)
        Write-Info (' Renamed files are in: {0}' -f $filesFull)
        if ($attentionFull) { Write-Info (' Files needing attention are in: {0}' -f $attentionFull) 'Yellow' }
        if ($zipFull) { Write-Info (' New zip:              {0}' -f $zipFull) }
    }
    Write-Info ''

    return [pscustomobject]@{
        Mode            = $(if ($Apply) { 'Apply' } else { 'DryRun' })
        Source          = $srcFull
        DestinationPrefix = $prefix
        Budget          = $budget
        ReportPath      = $reportFull
        LogPath         = $logFull
        FilesFolder     = $filesFull
        AttentionFolder = $attentionFull
        ZipPath         = $zipFull
        FilesProcessed  = $processed
        FilesRenamed    = $renamed
        FilesFlagged    = $flagged
        ItemsSkipped    = $skipped
        ItemsRejected   = $rejected
        Errors          = $errors
        Rows            = $rows
        LogRows         = $ctx.LogRows.ToArray()
        OutputIsSynced  = $outputIsSynced
        Warnings        = $ctx.Warnings.ToArray()
    }
}

Export-ModuleMember -Function Invoke-PathShortener, Get-SyncedFolderMatch
