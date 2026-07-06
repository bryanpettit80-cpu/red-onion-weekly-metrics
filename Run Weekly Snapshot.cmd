@echo off
setlocal
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_program\Run-WeeklySnapshot.ps1"

echo.
pause
