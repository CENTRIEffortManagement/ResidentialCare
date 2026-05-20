[CmdletBinding()]
param(
    [switch] $ValidateOnly
)

$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot '..'))
$sharedRefresh = [System.IO.Path]::GetFullPath((Join-Path $projectRoot '..\mcode-pq-tools\scripts\refresh.ps1'))
$configPath = Join-Path $projectRoot 'pq.project.json'
$configRelativePath = [System.IO.Path]::GetRelativePath($projectRoot, $configPath)

$parameters = @{
    ProjectRoot = $projectRoot
    Config = $configRelativePath
}

if ($ValidateOnly) {
    $parameters.ValidateOnly = $true
}

& $sharedRefresh @parameters
exit $LASTEXITCODE
