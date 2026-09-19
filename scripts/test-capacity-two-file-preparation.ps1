param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$canonicalPath = Join-Path $repoRoot 'Workflows/ResidentialCare/CapacityDistribution/Capacity.pq'
$currentPath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity.xlsx_PowerQuery.m'
$script:checks = 0

function Assert-CapacityPreparation {
    param([bool] $Condition, [string] $Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:checks++
}

function Get-SharedQueryNames {
    param([string] $Source)
    $pattern = '(?m)^shared\s+(?<name>#"[^"]+"|[A-Za-z0-9_]+)\s*='
    return @([regex]::Matches($Source, $pattern) | ForEach-Object { $_.Groups['name'].Value })
}

$canonical = [IO.File]::ReadAllText($canonicalPath)
$current = [IO.File]::ReadAllText($currentPath)
$canonicalNames = @(Get-SharedQueryNames $canonical)
$currentNames = @(Get-SharedQueryNames $current)

Assert-CapacityPreparation (($canonicalNames -join "`n") -eq ($currentNames -join "`n")) 'prepared source preserves every existing shared query interface in order'
Assert-CapacityPreparation ([regex]::Matches($canonical, 'CSharpWorkbookName\s*=\s*if Comparer\.OrdinalIgnoreCase\(SemanticRole, "AIN"\)\s*=\s*0 then "CapacityDistrib\(B\)-shifts\.xlsx" else "CapacityDistrib\(A\.2\)-shifts\.xlsx"').Count -eq 3) 'all three C# imports route by semantic role value'
Assert-CapacityPreparation ([regex]::Matches($canonical, 'Source\s*=\s*Excel\.Workbook\(File\.Contents\(Folder\s*&\s*"\\"\s*&\s*SemanticRole\s*&\s*"\\"\s*&\s*CSharpWorkbookName\)').Count -eq 3) 'all three C# imports use the semantic route result'
Assert-CapacityPreparation ([regex]::Matches($canonical, 'Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"').Count -eq 3) 'C# compatibility table contract is unchanged for every role slot'
Assert-CapacityPreparation ($canonical.Contains('and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")')) 'resolver uses a folder-boundary-safe prefix match'
Assert-CapacityPreparation ($canonical.Contains('Comparer.OrdinalIgnoreCase(InputFileName, "Capacity.xlsx") = 0')) 'resolver validates the intended workbook filename'
Assert-CapacityPreparation (-not ($canonical -match 'C:\\Users\\(?!Public\\)[^"\\]+\\')) 'prepared source contains no user-specific local path'
Assert-CapacityPreparation ($canonical.Contains('Prepared only. Do not synchronize')) 'prepared source records its no-sync gate'

Write-Host "PASS: $script:checks Capacity preparation checks. No workbook opened."
