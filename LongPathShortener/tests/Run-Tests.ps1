<#
    Run-Tests.ps1

    Builds the synthetic fixtures and runs the Pester tests.

    On Windows 10 or 11, from the LongPathShortener folder:
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1

    Uses Pester 3.x, which ships with Windows PowerShell 5.1. If a newer Pester
    is also installed, version 3 is still picked. -PesterPath points at a
    specific copy of Pester if needed.

    Exit code is the number of failed tests (0 means everything passed).
#>
[CmdletBinding()]
param(
    [string]$PesterPath = ''
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# Pester 3 puts its temporary TestDrive under $env:TEMP.
if (-not $env:TEMP) { $env:TEMP = [System.IO.Path]::GetTempPath() }

Get-Module Pester | Remove-Module -Force
if ($PesterPath) {
    Import-Module $PesterPath -Force
} else {
    $pester = Get-Module -ListAvailable -Name Pester | Where-Object { $_.Version.Major -eq 3 } | Sort-Object -Property Version -Descending | Select-Object -First 1
    if (-not $pester) {
        throw 'Pester 3 was not found. It is built into Windows PowerShell 5.1 (C:\Program Files\WindowsPowerShell\Modules\Pester).'
    }
    Import-Module $pester.Path -Force
}

Write-Host ('PowerShell {0} on {1}, Pester {2}' -f $PSVersionTable.PSVersion, [System.Environment]::OSVersion.VersionString, (Get-Module Pester).Version)

$result = Invoke-Pester -Script (Join-Path $here 'LongPathShortener.Tests.ps1') -PassThru
Write-Host ''
Write-Host ('Passed: {0}  Failed: {1}  Total: {2}' -f $result.PassedCount, $result.FailedCount, $result.TotalCount)
exit $result.FailedCount
