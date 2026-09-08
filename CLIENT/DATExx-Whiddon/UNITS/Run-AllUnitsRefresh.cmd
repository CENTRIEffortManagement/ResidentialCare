@echo off
setlocal

set "RUN_ROOT=%~dp0"
set "ALL_UNITS_SCRIPT=%RUN_ROOT%runner\Invoke-AllUnitsRefresh.ps1"

if not exist "%ALL_UNITS_SCRIPT%" (
    echo Required ResidentialCare All Units Refresh runner script is missing:
    echo %ALL_UNITS_SCRIPT%
    exit /b 10
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%ALL_UNITS_SCRIPT%" %*
set "ALL_UNITS_EXIT_CODE=%ERRORLEVEL%"

exit /b %ALL_UNITS_EXIT_CODE%
