@echo off
setlocal
title Red Onion Weekly Snapshot
set "OPERATIONS_ROOT=%~dp0."
set "PROGRAM_ROOT=%~dp0Red Onion Weekly Metrics Automation"
set "OUTER_WORKSPACE=1"

if not exist "%PROGRAM_ROOT%\_program\Run-WeeklySnapshot.ps1" (
  set "PROGRAM_ROOT=%~dp0."
  set "OUTER_WORKSPACE=0"
)

set "RUNNER=%PROGRAM_ROOT%\_program\Run-WeeklySnapshot.ps1"

echo.
echo RED ONION WEEKLY SNAPSHOT
echo =========================
echo The program will process the current Tuesday-Sunday reports,
echo create the finished workbooks, and archive successful inputs.
echo.

if not exist "%RUNNER%" (
  echo ATTENTION NEEDED: The automation program was not found:
  echo %RUNNER%
  echo Ask for technical help before moving any files.
  echo.
  pause
  exit /b 2
)

if "%OUTER_WORKSPACE%"=="1" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%" -OperationsRoot "%OPERATIONS_ROOT%" %*
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%" %*
)
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
  echo Finished. Review the workbook paths shown above.
) else (
  echo ATTENTION NEEDED: The weekly snapshot did not finish normally.
  echo The current reports were left in place for review.
)
echo.
echo You may close this window after reviewing the result.
pause
exit /b %EXITCODE%
