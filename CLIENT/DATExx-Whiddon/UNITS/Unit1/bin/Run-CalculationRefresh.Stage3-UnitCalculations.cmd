@echo off
setlocal

set "RUN_ROOT=%~dp0"
set "CALCULATION_SCRIPT=%RUN_ROOT%runner\Invoke-CalculationRefresh.ps1"

if not exist "%CALCULATION_SCRIPT%" (
    echo Required ResidentialCare Unit Calculations runner script is missing:
    echo %CALCULATION_SCRIPT%
    exit /b 10
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%CALCULATION_SCRIPT%" %*
set "CALCULATION_EXIT_CODE=%ERRORLEVEL%"

exit /b %CALCULATION_EXIT_CODE%
