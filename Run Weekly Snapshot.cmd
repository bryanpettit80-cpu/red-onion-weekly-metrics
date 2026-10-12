@echo off
setlocal
title Red Onion Weekly Snapshot
set "OPERATIONS_ROOT=%~dp0."
set "RUNNER=%~dp0_program\Run-WeeklySnapshot.ps1"
set "PAUSE_AT_END=1"
set "EXPLICIT_ROOT=0"
for %%A in (%*) do (
  if /I "%%~A"=="-NonInteractive" set "PAUSE_AT_END=0"
  if /I "%%~A"=="-OperationsRoot" set "EXPLICIT_ROOT=1"
)

if exist "%RUNNER%" goto run_standalone

set "PROGRAM_COUNT=0"
for /d %%D in ("%~dp0*") do if exist "%%~fD\_program\Run-WeeklySnapshot.ps1" (
  set /a PROGRAM_COUNT+=1 >nul
  set "RUNNER=%%~fD\_program\Run-WeeklySnapshot.ps1"
)
if not "%PROGRAM_COUNT%"=="1" (
  echo ATTENTION NEEDED: Expected one program folder beside the operator folders.
  echo Keep one installed program here, or launch its own Run Weekly Snapshot.cmd.
  set "EXITCODE=2"
  goto finish
)
if "%EXPLICIT_ROOT%"=="1" goto run_standalone
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%" -OperationsRoot "%OPERATIONS_ROOT%" %*
set "EXITCODE=%ERRORLEVEL%"
goto finish

:run_standalone
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%" %*
set "EXITCODE=%ERRORLEVEL%"

:finish
echo.
if "%EXITCODE%"=="0" (
  echo Finished. Review the run outcome and workbook paths shown above.
) else (
  echo ATTENTION NEEDED: Review the outcome above and LAST RUN STATUS before retrying.
  echo A report may already have been published; use its recorded outcome.
)
if "%PAUSE_AT_END%"=="1" (
  echo You may close this window after reviewing the result.
  pause
)
exit /b %EXITCODE%
