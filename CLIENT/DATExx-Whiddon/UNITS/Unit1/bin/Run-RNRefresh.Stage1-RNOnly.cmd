@echo off
setlocal

set "RUN_ROOT=%~dp0"
set "RN_SCRIPT=%RUN_ROOT%runner\Invoke-RNRefresh.ps1"

if not exist "%RN_SCRIPT%" (
    echo Required ResidentialCare RN runner script is missing:
    echo %RN_SCRIPT%
    exit /b 10
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%RN_SCRIPT%" %*
set "RN_EXIT_CODE=%ERRORLEVEL%"

exit /b %RN_EXIT_CODE%
