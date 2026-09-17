@echo off
rem Starts PSDocling from this checkout: hidden PowerShell, app window.
start "" /B powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0app\scripts\Start-PSDocling.ps1" %*
exit /b 0
