@echo off
setlocal

set "RUN_ROOT=%~dp0"
set "ROLE_SCRIPT=%RUN_ROOT%runner\Invoke-RoleRefresh.ps1"

if not exist "%ROLE_SCRIPT%" (
    echo Required ResidentialCare Roles runner script is missing:
    echo %ROLE_SCRIPT%
    exit /b 10
)

pwsh -NoProfile -ExecutionPolicy Bypass -File "%ROLE_SCRIPT%" %*
set "ROLE_EXIT_CODE=%ERRORLEVEL%"

exit /b %ROLE_EXIT_CODE%
