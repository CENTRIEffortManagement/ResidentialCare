@echo off
setlocal

set "RUN_ROOT=%~dp0"
set "UNIT_SCRIPT=%RUN_ROOT%runner\Invoke-UnitRefresh.ps1"

if not exist "%UNIT_SCRIPT%" (
    echo Required ResidentialCare Unit Refresh runner script is missing:
    echo %UNIT_SCRIPT%
    exit /b 10
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%UNIT_SCRIPT%" %*
set "UNIT_EXIT_CODE=%ERRORLEVEL%"

exit /b %UNIT_EXIT_CODE%
