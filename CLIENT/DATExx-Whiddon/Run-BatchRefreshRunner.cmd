@echo off
setlocal
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0runner\Invoke-BatchRefreshRunner.ps1" %*
set "BATCH_RUNNER_EXIT=%ERRORLEVEL%"
rem Command-line callers must retain their exit code and never wait for a key.
if not "%~1"=="" exit /b %BATCH_RUNNER_EXIT%
rem Successful menu results already wait in PowerShell; explicit Exit does not.
if "%BATCH_RUNNER_EXIT%"=="0" exit /b 0
echo.
echo The runner stopped. Please read the explanation above.
echo Press any key to close this window.
pause >nul
exit /b %BATCH_RUNNER_EXIT%
