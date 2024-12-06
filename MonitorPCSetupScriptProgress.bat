@echo off
:: Define the path to the PowerShell script
set SCRIPT_PATH="C:\IT\MonitorProgressInEventLog.ps1"

:: Run the PowerShell script with Unrestricted execution policy
powershell.exe -NoProfile -ExecutionPolicy Unrestricted -File "C:\IT\MonitorProgressInEventLog.ps1"

:: Pause for user feedback (optional)
pause
