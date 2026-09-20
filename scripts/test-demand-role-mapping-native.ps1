[CmdletBinding()]
param(
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-AIN-B-PQTest-2.155.2/package/tools/PQTest.exe')
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$sourcePath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/2-DemandExtract.xlsx_PowerQuery.m'
$fixturePath = Join-Path $repoRoot 'Workflows/ResidentialCare/Diagnostics/DemandRoleMapping_TEST.pq'

function Get-MDefinitions {
    param([string] $Text)
    foreach ($match in [regex]::Matches($Text, '(?m)^shared\s+(?<name>#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=')) {
        $start = $match.Index + $match.Length
        $inString = $false
        $lineComment = $false
        $end = -1
        for ($index = $start; $index -lt $Text.Length; $index++) {
            $current = $Text[$index]
            $next = if ($index + 1 -lt $Text.Length) { $Text[$index + 1] } else { [char]0 }
            if ($lineComment) {
                if ($current -eq "`n") { $lineComment = $false }
                continue
            }
            if ($inString) {
                if ($current -eq '"' -and $next -eq '"') { $index++ }
                elseif ($current -eq '"') { $inString = $false }
                continue
            }
            if ($current -eq '/' -and $next -eq '/') { $lineComment = $true; $index++; continue }
            if ($current -eq '"') { $inString = $true; continue }
            if ($current -eq ';') { $end = $index; break }
        }
        if ($end -lt 0) { throw "Unterminated M definition: $($match.Groups['name'].Value)" }
        $token = $match.Groups['name'].Value
        $name = if ($token.StartsWith('#"')) { $token.Substring(2, $token.Length - 3).Replace('""', '"') } else { $token }
        [pscustomobject]@{ Name = $name; Token = $token; Expression = $Text.Substring($start, $end - $start).Trim() }
    }
}

$source = [IO.File]::ReadAllText($sourcePath)
$definitions = @(Get-MDefinitions $source)
if ($definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) { throw 'Duplicate shared query names.' }
$overrides = @{
    'IMPORT CentriSyncPaths' = 'Inputs[PathNavigation]'
    'IMPORT Distributed FTE Workbook' = 'error "Workbook access is disabled in native fixtures."'
    'IMPORT Distributed FTE Allocation' = 'Inputs[Allocation]'
    'IMPORT Distributed FTE Profile' = 'Inputs[Profile]'
    'IMPORT Distributed FTE Checks' = 'Inputs[Checks]'
    'IMPORT Settings Data' = 'Inputs[Settings]'
    'IMPORT Role Lists' = 'error "SharePoint access is disabled in native fixtures."'
    'LINK Roles' = 'Inputs[Roles]'
    'LINK RoleAnalysis' = 'Inputs[Analysis]'
}
foreach ($name in $overrides.Keys) {
    if ($definitions.Name -notcontains $name) { throw "Missing fixture override target: $name" }
}
$bindings = foreach ($definition in $definitions) {
    $expression = $definition.Expression
    if ($overrides.ContainsKey($definition.Name)) { $expression = $overrides[$definition.Name] }
    if ($definition.Name -eq 'UnitL1PathTABLE') {
        $lookup = 'Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content]'
        if (-not $expression.Contains($lookup)) { throw 'Cannot replace the single FilePathUrl input in the native fixture.' }
        $expression = $expression.Replace($lookup, 'Inputs[PathInput]')
    }
    $definition.Token + ' = ' + $expression
}
$evaluate = '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + @'

in [Demand=ShiftUnitDemandHRS, Checks=DemandExtraction_CHECK, Roles=#"Demand Roles",
    Pattern=#"Distributed FTE Prepare", SourceCells=#"Distributed FTE Source Cells",
    Exclusions=#"Distributed FTE Exclusions", Reconciliation=DemandExtraction_RECONCILIATION,
    RoleCode=#"Demand Role Code", UnitPath=Unit1Path, CalendarCheck=#"Demand Calendar CHECK",
    Calendar=#"Demand Calendar Prepare"]
'@
$fixture = @(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)) | Where-Object Name -eq 'DemandRoleFixtures')
if ($fixture.Count -ne 1) { throw 'Expected one DemandRoleFixtures function.' }
$testSource = "section DemandRoleNative;`nshared DemandRoleTests = () as table => let`nEvaluate = " +
    $evaluate + ",`nFixtures = " + $fixture[0].Expression + "`nin Fixtures(Evaluate);`n"
if ($testSource -match 'File\.Contents\s*\(|Excel\.(Workbook|CurrentWorkbook)\s*\(|SharePoint\.Tables\s*\(') {
    throw 'Native tests must not contain file, workbook or SharePoint access.'
}
if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) { throw 'Microsoft Power Query SDK Tools 2.155.2 are required.' }

# PQTest needs the signed-in Windows profile even though these fixtures have no data sources.
# Probe before starting the executable, which otherwise may crash during DPAPI initialization.
$probe = [Text.Encoding]::UTF8.GetBytes('ResidentialCare-DemandRole-PQTest-Probe')
$encrypted = [Security.Cryptography.ProtectedData]::Protect($probe, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
$roundTrip = [Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
if (-not [Linq.Enumerable]::SequenceEqual[byte]($probe, $roundTrip)) { throw 'PQTest profile probe failed.' }

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('ResidentialCare-DemandRole-Tests-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temporaryRoot
$runnerProcess = $null
try {
    $extensionPath = Join-Path $temporaryRoot 'DemandRoleTests.mez'
    $queryPath = Join-Path $temporaryRoot 'fixtures.query.pq'
    $resultPath = Join-Path $temporaryRoot 'result.json'
    $errorPath = Join-Path $temporaryRoot 'stderr.txt'
    $credentialsPath = Join-Path $temporaryRoot 'credentials.bin'
    $archive = [IO.Compression.ZipFile]::Open($extensionPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entry = $archive.CreateEntry('DemandRoleTests.pq')
        $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write($testSource) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
    [IO.File]::WriteAllText($queryPath, 'DemandRoleTests()', [Text.UTF8Encoding]::new($false))
    $arguments = '-cfp "' + $credentialsPath + '" run-test --extension "' + $extensionPath +
        '" --queryFile "' + $queryPath + '" --prettyPrint'
    $runnerProcess = Start-Process -FilePath $PQTestPath -ArgumentList $arguments -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput $resultPath -RedirectStandardError $errorPath
    if (-not $runnerProcess.WaitForExit(60000)) {
        $runnerProcess.Kill($true)
        throw 'Power Query fixture execution exceeded 60 seconds.'
    }
    $lines = @(Get-Content -LiteralPath $resultPath)
    $jsonStart = -1
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if ($lines[$index].TrimStart().StartsWith('[') -or $lines[$index].TrimStart().StartsWith('{')) { $jsonStart = $index; break }
    }
    if ($jsonStart -lt 0) { throw "No Power Query JSON result: $($lines -join ' ') $([IO.File]::ReadAllText($errorPath))" }
    $results = @(($lines[$jsonStart..($lines.Count - 1)] -join "`n") | ConvertFrom-Json)
    $failures = @($results | Where-Object { $_.Status -ne 'Passed' })
    $failedRows = @($results | ForEach-Object { @($_.Output | Where-Object { $_.Passed -ne $true }) })
    if ($runnerProcess.ExitCode -ne 0 -or $results.Count -eq 0 -or $failures.Count -gt 0 -or $failedRows.Count -gt 0) {
        $results | ConvertTo-Json -Depth 15
        throw 'Demand role native fixtures failed.'
    }
    [ordered]@{
        Status = 'Passed'
        FixtureRows = [int]($results | Measure-Object -Property RowCount -Sum).Sum
        FailedRows = 0
        DataSources = @($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
        Scope = 'Actual Demand Extract M queries with synthetic inputs; no workbook or SharePoint access'
    } | ConvertTo-Json
}
finally {
    if ($null -ne $runnerProcess) {
        if (-not $runnerProcess.HasExited) { $runnerProcess.Kill($true) }
        $runnerProcess.Dispose()
    }
    # Only remove this invocation's verified temporary directory.
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $allowedParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolvedTemporaryRoot.StartsWith($allowedParent, [StringComparison]::OrdinalIgnoreCase) -or
        -not ([IO.Path]::GetFileName($resolvedTemporaryRoot)).StartsWith('ResidentialCare-DemandRole-Tests-')) {
        throw 'Unexpected test temporary directory; cleanup refused.'
    }
    Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
}
