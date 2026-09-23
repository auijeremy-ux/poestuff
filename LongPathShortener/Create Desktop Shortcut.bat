@echo off
rem Puts a "Long Path Shortener" shortcut on your desktop. Run it once.
rem You can drag zip files and folders onto the shortcut.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0engine\New-DesktopShortcut.ps1"
pause
