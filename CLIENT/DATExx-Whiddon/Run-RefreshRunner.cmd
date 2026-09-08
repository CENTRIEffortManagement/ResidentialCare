@echo off
setlocal

set "DATE_ROOT=%~dp0"
set "REFRESH_SCRIPT=%DATE_ROOT%runner\Invoke-RefreshRunner.ps1"

if not exist "%REFRESH_SCRIPT%" (
    echo Required ResidentialCare Date Refresh runner script is missing:
    echo %REFRESH_SCRIPT%
    exit /b 10
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%REFRESH_SCRIPT%" %*
set "REFRESH_EXIT_CODE=%ERRORLEVEL%"

exit /b %REFRESH_EXIT_CODE%
