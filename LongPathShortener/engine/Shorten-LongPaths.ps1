<#
.SYNOPSIS
    Finds paths that will be too long once files sit in a synced SharePoint
    folder, and copies or extracts them with shorter names.

.DESCRIPTION
    Works on a folder or a .zip file. By default it is a dry run: it writes a
    CSV report to the output folder and changes nothing. Add -Apply to copy the
    folder, or extract the zip, into the output folder with the new names.

    The source folder or zip is never changed. File contents are never changed.
    Runs offline in Windows PowerShell 5.1 with nothing to install.

.PARAMETER Source
    The folder or .zip file to check.

.PARAMETER DestinationPrefix
    The full local path of the synced SharePoint folder the files will finally
    live in, for example C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd.
    Path lengths are worked out as if the files were already there.

.PARAMETER OutputFolder
    Where the report, the log and the renamed files go. Default C:\CL\Out.

.PARAMETER MaxPathLength
    The longest full path allowed. Default 218, which is Excel's limit.

.PARAMETER SafetyMargin
    Characters kept spare below MaxPathLength. Default 10.

.PARAMETER AbbreviationsCsv
    Optional CSV file with columns Find and Replace.

.PARAMETER Apply
    Actually copy or extract the files. Without this switch nothing is written
    except the report.

.PARAMETER ExpandNestedZips
    Also unpack zips found inside the source, up to 3 levels deep.

.PARAMETER CreateZip
    With -Apply, also make a new zip of the renamed files.

.PARAMETER MaxTotalSizeGB
    Stop if the zip contents add up to more than this. Default 20.

.PARAMETER AllowSyncedOutput
    Skip the question asked when the output folder is inside OneDrive or a
    synced SharePoint library.

.PARAMETER IncludeSystemFiles
    Keep __MACOSX folders, .DS_Store, Thumbs.db and desktop.ini files instead
    of skipping them.

.EXAMPLE
    .\Shorten-LongPaths.ps1 -Source 'C:\Users\jsmith\Downloads\Smith docs.zip' -DestinationPrefix 'C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd'

    Dry run. Writes a report to C:\CL\Out and changes nothing.

.EXAMPLE
    .\Shorten-LongPaths.ps1 -Source 'C:\Users\jsmith\Downloads\Smith docs.zip' -DestinationPrefix 'C:\Users\jsmith\CL\Matters - Documents\Smith Pty Ltd' -AbbreviationsCsv ..\Abbreviations.csv -ExpandNestedZips -Apply

    Extracts the zip into C:\CL\Out\Smith docs with shortened names and writes
    the evidence log.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Source,
    [Parameter(Mandatory = $true)][string]$DestinationPrefix,
    [string]$OutputFolder = 'C:\CL\Out',
    [int]$MaxPathLength = 218,
    [int]$SafetyMargin = 10,
    [string]$AbbreviationsCsv = '',
    [switch]$Apply,
    [switch]$ExpandNestedZips,
    [switch]$CreateZip,
    [double]$MaxTotalSizeGB = 20,
    [switch]$AllowSyncedOutput,
    [switch]$IncludeSystemFiles
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'LongPathShortener.psm1') -Force

$params = @{}
foreach ($key in $PSBoundParameters.Keys) { $params[$key] = $PSBoundParameters[$key] }
if (-not $params.ContainsKey('OutputFolder')) { $params.OutputFolder = $OutputFolder }

try {
    $null = Invoke-PathShortener @params
    exit 0
} catch {
    Write-Host ''
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
