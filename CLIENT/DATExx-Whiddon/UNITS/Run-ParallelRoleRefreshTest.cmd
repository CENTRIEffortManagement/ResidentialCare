@echo off
setlocal

set "RUN_ROOT=%~dp0"
set "PARALLEL_SCRIPT=%RUN_ROOT%parallel-test\Invoke-ParallelRoleRefreshTest.ps1"

if not exist "%PARALLEL_SCRIPT%" (
    echo Required parallel role refresh test script is missing:
    echo %PARALLEL_SCRIPT%
    exit /b 10
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%PARALLEL_SCRIPT%" %*
exit /b %ERRORLEVEL%
