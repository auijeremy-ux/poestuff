@echo off
rem -----------------------------------------------------------------------
rem  Long path shortener
rem  Drag a zip file or a folder onto this file, or double-click it.
rem
rem  Runs the built-in Windows PowerShell with the execution policy bypassed
rem  for this one process only. Nothing is installed and nothing goes online.
rem -----------------------------------------------------------------------
setlocal
title Long path shortener
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0engine\Start-Interactive.ps1" %1
set "RC=%ERRORLEVEL%"
rem 0 and 2 are handled by the script itself. Anything else means PowerShell
rem could not start the script, so keep the window open to show why.
if not "%RC%"=="0" if not "%RC%"=="2" pause
endlocal & exit /b %RC%
