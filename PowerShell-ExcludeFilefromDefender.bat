@setlocal enableextensions
@cd /d "%~dp0"

PowerShell.exe -ExecutionPolicy Bypass -File PowerShell-ExcludeFilefromDefender.ps1
pause