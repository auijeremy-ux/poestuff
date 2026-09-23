<#
    New-TestFixtures.ps1

    Builds synthetic test data. No real client data is used anywhere.

    Creates, under -Path:
      SourceFolder\   a folder tree 12 levels deep with names over 100 characters,
                      reserved names, trailing periods and spaces, extra spaces,
                      dated and numbered file names, names that collide once
                      shortened, housekeeping files, an empty folder and a zip
                      inside the folder.
      Source.zip      the same tree plus illegal characters, names that differ
                      only in case, zip slip entries, an entry marked as
                      encrypted, zips nested 4 levels deep, Mac clutter, and
                      accented names stored the Windows way and the Mac way.

    Returns an object describing what was created.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Path
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$isWin = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
$sep = [string][System.IO.Path]::DirectorySeparatorChar
$rng = New-Object System.Random -ArgumentList 20260923

function ConvertTo-FixtureLongPath {
    param([string]$P)
    if (-not $isWin) { return $P }
    if ($P.StartsWith('\\?\')) { return $P }
    if ($P.StartsWith('\\')) { return '\\?\UNC\' + $P.Substring(2) }
    return '\\?\' + $P
}

function New-RandomBytes {
    param([int]$Size)
    $b = New-Object byte[] $Size
    $rng.NextBytes($b)
    return , $b
}

$script:dateCounter = 0
function Get-NextDate {
    # Even seconds, because zip files only store time to 2 second precision.
    $script:dateCounter++
    return (New-Object DateTime -ArgumentList 2021, 1, 1, 10, 20, 30).AddDays($script:dateCounter * 3)
}

# ---------------------------------------------------------------------------
# The tree shared by the folder and the zip
# ---------------------------------------------------------------------------

$deep = @(
    'Smith Pty Ltd - Client Documents Supplied by the Client for the Construction Dispute with Jones Building Pty Ltd',
    'Correspondence and Construction Records for the Period of the Project from Commencement to Practical Completion',
    'Level 03 - Variations and Payment Claims Submitted by the Builder for the Construction of the Residential Units',
    'Level 04 - Correspondence between the Owner and the Builder regarding the Defective Waterproofing of the Bathrooms',
    'Level 05 - Expert Reports and Photographs of the Defects Prepared for the Adjudication Application and Response',
    'Level 06 - Subcontractor Agreements and Specifications for the Plumbing and Electrical Works on the Upper Levels',
    'Level 07 - Minutes of the Site Meetings Held between the Superintendent and the Contractor During Construction',
    'Level 08 - Extension of Time Claims and Notices of Delay Issued under the Contract by the Builder and Consultants',
    'Level 09 - Documents Produced under Subpoena by the Certifier and the Council for the Tribunal Hearing in March',
    'Level 10 - Draft Affidavits and Statements of Evidence Prepared for the Owner and the Independent Expert Witness',
    'Level 11 - Invoices and Receipts for the Rectification Works Carried Out by Other Trades after the Termination',
    'Level 12 - Final Bundle of the Documents for the Hearing Including the Index and the Chronology of the Key Events'
)

function Get-DeepPath {
    param([int]$Levels)
    return , [string[]]($deep[0..($Levels - 1)])
}

$longLetter = 'Letter from the builder to the owner regarding the defective waterproofing to the upstairs bathrooms and the laundry, with the photographs, the invoices and the expert commentary attached, version '

$common = New-Object System.Collections.Generic.List[object]
function Add-Common {
    param([string[]]$Segments, [int]$Size = 2048)
    $common.Add([pscustomobject]@{ Segments = $Segments; Bytes = (New-RandomBytes $Size); Date = (Get-NextDate) })
}

Add-Common (@($deep[0]) + 'Index of documents.txt')
Add-Common (@($deep[0]) + 'Thumbs.db') 64
Add-Common ((Get-DeepPath 2) + 'Scope of works.pdf')
Add-Common ((Get-DeepPath 3) + '2025-11-02 Site inspection report prepared by the independent building consultant engaged by the owner.pdf')
Add-Common ((Get-DeepPath 6) + '003 Photographs of the defective waterproofing membrane taken during the site inspection.pdf') 8192
$d12 = Get-DeepPath 12
Add-Common ($d12 + '2026-03-14 Letter from the builder to the owner regarding the defective waterproofing to the bathrooms version A.pdf')
Add-Common ($d12 + '2026-03-14 Letter from the builder to the owner regarding the defective waterproofing to the bathrooms version B.pdf')
Add-Common ($d12 + '001 Statutory declaration of the site supervisor regarding the progress of the construction works.docx')
Add-Common ($d12 + 'Payment claim schedule and supporting calculations for the variations to the contract works.xlsx') 4096
Add-Common ($d12 + 'CON.txt') 100
Add-Common ($d12 + 'aux.docx') 100
Add-Common @('Letters', ($longLetter + 'A.pdf'))
Add-Common @('Letters', ($longLetter + 'B.pdf'))
Add-Common @('PRN', 'readme.txt') 100
Add-Common @('Old notes. ', 'file.txt') 100
Add-Common @('Draft   with   extra   spaces .docx') 100
Add-Common @('Notes about the meeting.') 100
Add-Common @('Smith Pty Ltd - Invoices', 'Smith Pty Ltd - Invoice 0001 for progress claim.pdf')
Add-Common @('Correspondence', '.DS_Store') 64

# ---------------------------------------------------------------------------
# Folder fixture
# ---------------------------------------------------------------------------

$root = [System.IO.Path]::GetFullPath($Path)
$folderRoot = $root + $sep + 'SourceFolder'
$zipPath = $root + $sep + 'Source.zip'

function Write-FixtureFile {
    param([string[]]$Segments, [byte[]]$Bytes, [datetime]$Date)
    $full = $folderRoot + $sep + ($Segments -join $sep)
    $dir = $full.Substring(0, $full.LastIndexOf($sep))
    [void][System.IO.Directory]::CreateDirectory((ConvertTo-FixtureLongPath $dir))
    $lp = ConvertTo-FixtureLongPath $full
    [System.IO.File]::WriteAllBytes($lp, $Bytes)
    [System.IO.File]::SetLastWriteTime($lp, $Date)
}

[void][System.IO.Directory]::CreateDirectory((ConvertTo-FixtureLongPath $folderRoot))
foreach ($f in $common) { Write-FixtureFile $f.Segments $f.Bytes $f.Date }
[void][System.IO.Directory]::CreateDirectory((ConvertTo-FixtureLongPath ($folderRoot + $sep + 'Empty folder')))

if (-not $isWin) {
    # Windows itself cannot hold these characters, so they only appear in the
    # folder fixture on other systems. The zip fixture always has them.
    Write-FixtureFile @('Invoice: 2024 "final" <v2>?.pdf') (New-RandomBytes 100) (Get-NextDate)
}

# ---------------------------------------------------------------------------
# Zip helpers
# ---------------------------------------------------------------------------

function New-ZipBytes {
    param($Entries)
    $ms = New-Object System.IO.MemoryStream
    $za = New-Object System.IO.Compression.ZipArchive -ArgumentList @($ms, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    foreach ($e in $Entries) {
        $ze = $za.CreateEntry($e.Name)
        $ze.LastWriteTime = New-Object System.DateTimeOffset -ArgumentList ([datetime]$e.Date)
        if ($null -ne $e.Bytes) {
            $s = $ze.Open()
            $s.Write($e.Bytes, 0, $e.Bytes.Length)
            $s.Dispose()
        }
    }
    $za.Dispose()
    return , $ms.ToArray()
}

function New-ZipEntrySpec {
    param([string]$Name, $Bytes, $Date)
    if ($null -eq $Date) { $Date = Get-NextDate }
    return [pscustomobject]@{ Name = $Name; Bytes = $Bytes; Date = $Date }
}

function Set-ZipEntryEncryptedFlag {
    # System.IO.Compression cannot write encrypted entries, so the test marks one
    # entry as encrypted by setting bit 0 of its flags in both of its headers.
    param([byte[]]$Bytes, [string]$EntryName)
    $latin = [System.Text.Encoding]::GetEncoding(28591)
    $text = $latin.GetString($Bytes)
    $needle = $latin.GetString([System.Text.Encoding]::UTF8.GetBytes($EntryName))
    $patched = 0
    $idx = $text.IndexOf($needle, 0, [System.StringComparison]::Ordinal)
    while ($idx -ge 0) {
        $l = $idx - 30
        $c = $idx - 46
        if ($l -ge 0 -and $Bytes[$l] -eq 0x50 -and $Bytes[$l + 1] -eq 0x4B -and $Bytes[$l + 2] -eq 3 -and $Bytes[$l + 3] -eq 4) {
            $Bytes[$l + 6] = $Bytes[$l + 6] -bor 1
            $patched++
        } elseif ($c -ge 0 -and $Bytes[$c] -eq 0x50 -and $Bytes[$c + 1] -eq 0x4B -and $Bytes[$c + 2] -eq 1 -and $Bytes[$c + 3] -eq 2) {
            $Bytes[$c + 8] = $Bytes[$c + 8] -bor 1
            $patched++
        }
        $idx = $text.IndexOf($needle, $idx + 1, [System.StringComparison]::Ordinal)
    }
    if ($patched -lt 2) { throw "Could not mark $EntryName as encrypted." }
}

function Set-ZipEntryNameBytes {
    # Swaps an ASCII placeholder name for raw bytes of the same length, in both
    # headers, without marking the name as UTF-8. Used to imitate zips made by
    # Windows Explorer (old DOS code page) and by Mac tools (unmarked UTF-8).
    param([byte[]]$Bytes, [string]$Placeholder, [byte[]]$NewName)
    if ($NewName.Length -ne $Placeholder.Length) { throw 'Placeholder and new name must be the same length.' }
    $latin = [System.Text.Encoding]::GetEncoding(28591)
    $text = $latin.GetString($Bytes)
    $patched = 0
    $idx = $text.IndexOf($Placeholder, 0, [System.StringComparison]::Ordinal)
    while ($idx -ge 0) {
        $l = $idx - 30
        $c = $idx - 46
        $isLocal = ($l -ge 0 -and $Bytes[$l] -eq 0x50 -and $Bytes[$l + 1] -eq 0x4B -and $Bytes[$l + 2] -eq 3 -and $Bytes[$l + 3] -eq 4)
        $isCentral = ($c -ge 0 -and $Bytes[$c] -eq 0x50 -and $Bytes[$c + 1] -eq 0x4B -and $Bytes[$c + 2] -eq 1 -and $Bytes[$c + 3] -eq 2)
        if ($isLocal -or $isCentral) {
            [System.Array]::Copy($NewName, 0, $Bytes, $idx, $NewName.Length)
            $patched++
        }
        $idx = $text.IndexOf($Placeholder, $idx + 1, [System.StringComparison]::Ordinal)
    }
    if ($patched -lt 2) { throw "Could not patch $Placeholder." }
}

# ---------------------------------------------------------------------------
# A zip inside the folder fixture
# ---------------------------------------------------------------------------

$innerForFolder = New-ZipBytes @(
    (New-ZipEntrySpec 'Inner letter.pdf' (New-RandomBytes 1500) $null),
    (New-ZipEntrySpec 'Sub folder/Inner note.txt' (New-RandomBytes 300) $null)
)
Write-FixtureFile @('Nested', 'Inner documents.zip') $innerForFolder (Get-NextDate)

# ---------------------------------------------------------------------------
# Zip fixture
# ---------------------------------------------------------------------------

$level4 = New-ZipBytes @((New-ZipEntrySpec 'L4 file.txt' (New-RandomBytes 200) $null))
$level3 = New-ZipBytes @((New-ZipEntrySpec 'L3 file.txt' (New-RandomBytes 200) $null), (New-ZipEntrySpec 'Level4.zip' $level4 $null))
$level2 = New-ZipBytes @((New-ZipEntrySpec 'L2 file.txt' (New-RandomBytes 200) $null), (New-ZipEntrySpec 'Level3.zip' $level3 $null))
$level1 = New-ZipBytes @((New-ZipEntrySpec 'L1 file.txt' (New-RandomBytes 200) $null), (New-ZipEntrySpec 'Level2.zip' $level2 $null), (New-ZipEntrySpec 'MXXller statement.txt' (New-RandomBytes 200) $null))
# "Mueller" with u-umlaut as unmarked UTF-8 (C3 BC), as Mac tools write it.
Set-ZipEntryNameBytes $level1 'MXXller statement.txt' ([byte[]](@(0x4D, 0xC3, 0xBC) + [System.Text.Encoding]::ASCII.GetBytes('ller statement.txt')))

$zipSlip = @('../../evil-1.txt', 'safe/../../evil-2.txt', '/tmp/evil-3.txt', 'C:/Windows/evil-4.txt', 'C:\evil-5.txt')
$encrypted = 'Confidential/Encrypted statement.pdf'

$specs = New-Object System.Collections.Generic.List[object]
foreach ($f in $common) { $specs.Add((New-ZipEntrySpec ($f.Segments -join '/') $f.Bytes $f.Date)) }
$specs.Add((New-ZipEntrySpec 'Empty folder/' $null $null))
$specs.Add((New-ZipEntrySpec 'Correspondence/Letter.pdf' (New-RandomBytes 900) $null))
$specs.Add((New-ZipEntrySpec 'Correspondence/CafX notes.txt' (New-RandomBytes 120) $null))
$specs.Add((New-ZipEntrySpec 'Correspondence/LETTER.pdf' (New-RandomBytes 901) $null))
$specs.Add((New-ZipEntrySpec 'Correspondence/Re: defects? "urgent" <draft>|final*.pdf' (New-RandomBytes 902) $null))
$specs.Add((New-ZipEntrySpec 'COM1/notes.txt' (New-RandomBytes 100) $null))
$specs.Add((New-ZipEntrySpec 'LPT1.docx' (New-RandomBytes 100) $null))
$specs.Add((New-ZipEntrySpec 'Windows style\sub folder\file.txt' (New-RandomBytes 100) $null))
$specs.Add((New-ZipEntrySpec $encrypted (New-RandomBytes 700) $null))
$specs.Add((New-ZipEntrySpec 'Nested/Level1.zip' $level1 $null))
$specs.Add((New-ZipEntrySpec '__MACOSX/Correspondence/._Letter.pdf' (New-RandomBytes 50) $null))
foreach ($z in $zipSlip) { $specs.Add((New-ZipEntrySpec $z (New-RandomBytes 50) $null)) }

$zipBytes = New-ZipBytes $specs
Set-ZipEntryEncryptedFlag $zipBytes $encrypted
# "Cafe" with e-acute in the DOS code page (0x82), as Windows Explorer writes it.
Set-ZipEntryNameBytes $zipBytes 'Correspondence/CafX notes.txt' ([byte[]]([System.Text.Encoding]::ASCII.GetBytes('Correspondence/Caf') + @(0x82) + [System.Text.Encoding]::ASCII.GetBytes(' notes.txt')))
[System.IO.File]::WriteAllBytes((ConvertTo-FixtureLongPath $zipPath), $zipBytes)

[pscustomobject]@{
    Root            = $root
    FolderPath      = $folderRoot
    ZipPath         = $zipPath
    ZipSlipEntries  = $zipSlip
    EncryptedEntry  = $encrypted
    DeepestLevels   = $deep.Count
    LongestName     = (@($deep) + @($common | ForEach-Object { $_.Segments[-1] }) | Measure-Object -Property Length -Maximum).Maximum
}
