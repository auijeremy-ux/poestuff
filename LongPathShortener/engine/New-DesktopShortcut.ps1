<#
    New-DesktopShortcut.ps1

    Puts a "Long Path Shortener" shortcut on the current user's desktop. The
    shortcut opens the app window without a console window, and zip files or
    folders can be dragged onto it.

    Run once, through "Create Desktop Shortcut.bat". Uses the built-in
    WScript.Shell object. Nothing is installed. Keep this file plain ASCII.
#>
$ErrorActionPreference = 'Stop'

try {
    $toolRoot = Split-Path -Parent $PSScriptRoot
    $app = Join-Path $PSScriptRoot 'Start-App.ps1'
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $desktop = [System.Environment]::GetFolderPath('Desktop')
    $link = Join-Path $desktop 'Long Path Shortener.lnk'

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($link)
    $shortcut.TargetPath = $powershell
    $shortcut.Arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}"' -f $app
    $shortcut.WorkingDirectory = $toolRoot
    $shortcut.Description = 'Shorten long file paths before putting client documents in SharePoint'
    # Built-in Windows folder icon.
    $shortcut.IconLocation = (Join-Path $env:SystemRoot 'System32\shell32.dll') + ',4'
    # Start minimised so the console never flashes up before the app hides it.
    $shortcut.WindowStyle = 7
    $shortcut.Save()

    Write-Host ''
    Write-Host '  Done. "Long Path Shortener" is now on your desktop.' -ForegroundColor Green
    Write-Host '  Double-click it to open the app, or drag a zip file or folder onto it.'
    Write-Host ''
    Write-Host '  If you move the LongPathShortener folder later, run this again.'
    Write-Host ''
} catch {
    Write-Host ''
    Write-Host ('  Could not create the shortcut: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host '  You can still open the app with "Long Path Shortener.bat".'
    Write-Host ''
    exit 1
}
