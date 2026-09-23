@echo off
rem -----------------------------------------------------------------------
rem  Long Path Shortener: opens the app window.
rem  You can also drag a zip file or folder onto this file.
rem
rem  Runs the built-in Windows PowerShell with the execution policy bypassed
rem  for this one process only. Nothing is installed and nothing goes online.
rem  If nothing appears, use "Shorten Long Paths (text version).bat", which
rem  shows any error on screen.
rem -----------------------------------------------------------------------
start "" powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0engine\Start-App.ps1" %1
