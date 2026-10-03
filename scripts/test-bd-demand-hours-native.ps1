[CmdletBinding()]
param(
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-AIN-B-PQTest-2.155.2/package/tools/PQTest.exe')
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$masterPath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m'
$extractPath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/2-DemandExtract.xlsx_PowerQuery.m'
$intervalPath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/DemandIntervals.xlsx_PowerQuery.m'
$fixturePath = Join-Path $repoRoot 'Workflows/ResidentialCare/Diagnostics/BDDemandHours_TEST.pq'

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

# Inject prepared history/targets and synthetic external navigation/settings snapshots.
# The changed Master distribution, hour checks and publication gates remain actual source code.
function New-NativeEvaluator {
    param([string] $Path, [hashtable] $Overrides, [string] $Output, [string] $Kind)
    $definitions = @(Get-MDefinitions ([IO.File]::ReadAllText($Path)))
    if ($definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) { throw "Duplicate query names in $Path" }
    foreach ($name in $Overrides.Keys) {
        if ($definitions.Name -notcontains $name) { throw "Missing $Kind fixture override target: $name" }
    }
    $bindings = foreach ($definition in $definitions) {
        $expression = $definition.Expression
        if ($Overrides.ContainsKey($definition.Name)) { $expression = $Overrides[$definition.Name] }
        if ($Kind -eq 'Extract' -and $definition.Name -eq 'UnitL1PathTABLE') {
            $lookup = 'Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content]'
            if (-not $expression.Contains($lookup)) { throw 'Cannot replace the exact FilePathUrl fixture input.' }
            $expression = $expression.Replace($lookup, 'Inputs[PathInput]')
        }
        if ($Kind -eq 'Master' -and $definition.Name -eq 'MW DayShift Allocation') {
            $expression = 'let NativeDistribution = (' + $expression +
                ') in if Record.HasFields(Inputs, "DistributionTransform") then Inputs[DistributionTransform](NativeDistribution) else NativeDistribution'
        }
        $definition.Token + ' = ' + $expression
    }
    '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + "`nin " + $Output
}

$disabled = 'error "External data access is disabled in native BD demand-hours fixtures."'
$masterOverrides = @{
    'IMPORT Master' = $disabled
    'INPUT SHiftEnd' = $disabled
    'INPUT MinuteWorkers' = $disabled
    'INPUT TargetMinutes' = $disabled
    'IMPORT Whiddon Lists' = $disabled
    'MW Historical WeekDayShift' = 'Inputs[History]'
    'MW Historical Weeks' = 'Table.Distinct(Table.SelectColumns(Inputs[History], {"Facility", "Week No", "FortnightWeek"}))'
    'MW TargetMinutes Prepare' = 'Inputs[Targets]'
    'MW TargetMinutes Source Prepare' = 'Inputs[SourceTargets]'
    'MW PreAllocation Check' = '#"MW Check Table"({})'
}
$masterOutput = '[Distribution=MinuteWorkersFTE_TABLE, Profile=MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE, ' +
    'Checks=MinuteWorkersFTE_CHECK, HourChecks=#"MW Fortnight Hours Check", ProfileChecks=#"MW FTE Profile Alignment Check", ' +
    'RawDistribution=#"MW DayShift Allocation", RawProfile=#"MW FTE Profile Comparison"]'
$extractOverrides = @{
    'IMPORT CentriSyncPaths' = 'Inputs[PathNavigation]'
    'IMPORT Distributed Demand' = 'Inputs[DemandNavigation]'
    'IMPORT Settings Data' = 'Inputs[Settings]'
    'IMPORT Role Lists' = $disabled
    'IMPORT Whiddon Facility Lists' = $disabled
    'LINK Roles' = 'Inputs[Roles]'
    'LINK RoleAnalysis' = 'Inputs[Analysis]'
    'LINK Facility Analysis' = '#table({"Title", "EffortManagementAnalysis", "LeaveBalanceAnalysis"}, {{"BD", true, false}})'
}
$extractOutput = '[Demand=ShiftUnitDemandHRS, Checks=DemandExtraction_CHECK, Roles=#"Demand Roles", ' +
    'Pattern=#"Distributed FTE Prepare", SourceCells=#"Distributed FTE Source Cells", ' +
    'Exclusions=#"Distributed FTE Exclusions", Reconciliation=DemandExtraction_RECONCILIATION]'
$intervalOverrides = @{
    'UnitL1PathTABLE' = $disabled
    'IMPORT Table_ShiftDemandHRS' = 'Table.RenameColumns(Inputs[Demand], {{"DurationOfShifts", "ShiftDuration"}})'
    'Table_IntervalIDList' = 'Inputs[IntervalIDs]'
    'Table_Intervals' = 'Inputs[Intervals]'
    'Meals' = '#table({"MealBreakstart", "MealBreakTime"}, {{5, 0.5}})'
}
$intervalOutput = '[Intervals=ShiftDemandUnitINTERVAL, Points=ShiftDemandINTERVALLIST, ' +
    'InputChecks=ShiftDemandHours_INPUT_CHECK, Checks=ShiftDemandIntervalHours_CHECK]'
$masterEvaluator = New-NativeEvaluator $masterPath $masterOverrides $masterOutput 'Master'
$extractEvaluator = New-NativeEvaluator $extractPath $extractOverrides $extractOutput 'Extract'
$intervalEvaluator = New-NativeEvaluator $intervalPath $intervalOverrides $intervalOutput 'Intervals'
$fixture = @(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)) | Where-Object Name -eq 'BDDemandHoursFixtures')
if ($fixture.Count -ne 1) { throw 'Expected one BDDemandHoursFixtures function.' }
$testSource = "section BDDemandHoursNative;`nshared BDDemandHoursTests = () as table => let`nEvaluateMaster = " +
    $masterEvaluator + ",`nEvaluateExtract = " + $extractEvaluator + ",`nEvaluateIntervals = " + $intervalEvaluator +
    ",`nFixtures = " + $fixture[0].Expression + "`nin Fixtures(EvaluateMaster, EvaluateExtract, EvaluateIntervals);`n"
# Fail closed even when a remaining connector call would be lazily unused.
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(|\bValue\.NativeQuery\s*\(') {
    throw 'Native BD tests must not contain file, workbook or network access.'
}
if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) { throw 'Microsoft Power Query SDK Tools 2.155.2 are required.' }

# PQTest needs the signed-in Windows profile even though these fixtures have no data sources.
# Probe before starting the executable, which otherwise may crash during DPAPI initialization.
$probe = [Text.Encoding]::UTF8.GetBytes('ResidentialCare-BDDemandHours-PQTest-Probe')
$encrypted = [Security.Cryptography.ProtectedData]::Protect($probe, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
$roundTrip = [Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
if (-not [Linq.Enumerable]::SequenceEqual[byte]($probe, $roundTrip)) { throw 'PQTest profile probe failed.' }

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('ResidentialCare-BDDemandHours-Tests-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temporaryRoot
$runnerProcess = $null
try {
    $extensionPath = Join-Path $temporaryRoot 'BDDemandHoursTests.mez'
    $queryPath = Join-Path $temporaryRoot 'fixtures.query.pq'
    $resultPath = Join-Path $temporaryRoot 'result.json'
    $errorPath = Join-Path $temporaryRoot 'stderr.txt'
    $credentialsPath = Join-Path $temporaryRoot 'credentials.bin'
    $archive = [IO.Compression.ZipFile]::Open($extensionPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entry = $archive.CreateEntry('BDDemandHoursTests.pq')
        $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write($testSource) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
    [IO.File]::WriteAllText($queryPath, 'BDDemandHoursTests()', [Text.UTF8Encoding]::new($false))
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
        throw 'BD demand hours native fixtures failed.'
    }
    [ordered]@{
        Status = 'Passed'
        FixtureRows = [int]($results | Measure-Object -Property RowCount -Sum).Sum
        FailedRows = 0
        DataSources = @($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
        Scope = 'Actual Master distribution stages, BD Extract and BD DemandIntervals with synthetic inputs; no external access'
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
        -not ([IO.Path]::GetFileName($resolvedTemporaryRoot)).StartsWith('ResidentialCare-BDDemandHours-Tests-')) {
        throw 'Unexpected test temporary directory; cleanup refused.'
    }
    Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
}
