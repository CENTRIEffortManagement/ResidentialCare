param([string] $Query = 'Table.Combine({CapacityDistribB_RunningTotals_CHECK, CapacityDistribB_InputKeys_CHECK})', [string] $ResultName = 'source-tests-result.json')
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$sourcePath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx_PowerQuery.m'
$source = [IO.File]::ReadAllText($sourcePath).Replace("`r`n", "`n")
$helperStart = $source.IndexOf('shared fnAddIndexedRunningTotal = ')
$helperEnd = $source.IndexOf('shared #"ResMaxReduction (C# Unassigned)" = ', $helperStart)
if ($helperStart -lt 0 -or $helperEnd -lt 0) { throw 'Cannot isolate the production running-total helper.' }
$helper = $source.Substring($helperStart, $helperEnd - $helperStart)
$fixtures = [IO.File]::ReadAllText((Join-Path $repoRoot 'Workflows/ResidentialCare/Diagnostics/CapacityDistribB_RunningTotals_TEST.pq'))
$fixtures = [regex]::Replace($fixtures, '^section [^;]+;', '')
$checkStart = $source.IndexOf('shared CapacityDistribB_INPUT_CHECK = ')
$checkEnd = $source.IndexOf('shared #"ResPeriodAvailabilityCapped(C#)SUM" = ', $checkStart)
if ($checkStart -lt 0 -or $checkEnd -lt 0) { throw 'Cannot isolate the production input check.' }
$checkDefinition = $source.Substring($checkStart, $checkEnd - $checkStart)
# Parameterize only the four source references; execute the production check logic unchanged.
$checkFixture = $checkDefinition.Replace('shared CapacityDistribB_INPUT_CHECK = ', 'shared fnCapacityDistribBTestInputCheck = (demandInput as table, allocationInput as table, cappedInput as table, priorityInput as table) as table => ')
$checkFixture = $checkFixture.Replace('CheckKeys("PeriodDemandTABLE", PeriodDemandTABLE,', 'CheckKeys("PeriodDemandTABLE", demandInput,')
$checkFixture = $checkFixture.Replace('CheckKeys("ResPeriodAllocationTABLE", ResPeriodAllocationTABLE,', 'CheckKeys("ResPeriodAllocationTABLE", allocationInput,')
$checkFixture = $checkFixture.Replace('CheckKeys("ResPeriodAvailabilityCapped(C#)TABLE", #"ResPeriodAvailabilityCapped(C#)TABLE",', 'CheckKeys("ResPeriodAvailabilityCapped(C#)TABLE", cappedInput,')
$checkFixture = $checkFixture.Replace('CheckKeys("IMPORT ResPeriodNWDTABLE", #"IMPORT ResPeriodNWDTABLE",', 'CheckKeys("IMPORT ResPeriodNWDTABLE", priorityInput,')
$priorityStart = $source.IndexOf('shared #"PrioritiseReductionAvail-Setup" = let')
$joinStart = $source.IndexOf('    #"Merged Queries" = Table.NestedJoin', $priorityStart)
$joinEnd = $source.IndexOf('    #"Inserted MINAVAILABLEREDUCTION"', $joinStart)
if ($priorityStart -lt 0 -or $joinStart -lt 0 -or $joinEnd -lt 0) { throw 'Cannot isolate the production priority join.' }
$joinSteps = $source.Substring($joinStart, $joinEnd - $joinStart).Trim().TrimEnd(',')
$joinSteps = $joinSteps.Replace('#"Expanded PeriodOverallocatedPriroristised"', 'leftInput').Replace('#"IMPORT ResPeriodNWDTABLE"', 'rightInput')
$joinFixture = "shared fnCapacityDistribBTestPriorityJoin = (leftInput as table, rightInput as table) as table => let`n$joinSteps`nin #`"Expanded IMPORT ResPeriodNWDTABLE`";`n"
$outputStart = $source.IndexOf('shared #"C###TABLE B" = let')
$gateStart = $source.IndexOf('    InputChecks = ', $outputStart)
$gateEnd = $source.IndexOf('    #"Merged Queries" = ', $gateStart)
if ($outputStart -lt 0 -or $gateStart -lt 0 -or $gateEnd -lt 0) { throw 'Cannot isolate the production output gate.' }
$gateSteps = $source.Substring($gateStart, $gateEnd - $gateStart).Trim().TrimEnd(',')
$gateSteps = $gateSteps.Replace('InputChecks = CapacityDistribB_INPUT_CHECK', 'InputChecks = checks').Replace('#"C##TABLE"', 'outputThunk()')
$gateFixture = "shared fnCapacityDistribBTestOutputGate = (checks as table, outputThunk as function) as table => let`n$gateSteps`nin Source;`n"
$inputFixtures = [IO.File]::ReadAllText((Join-Path $repoRoot 'Workflows/ResidentialCare/Diagnostics/CapacityDistribB_InputKeys_TEST.pq'))
$inputFixtures = [regex]::Replace($inputFixtures, '^section [^;]+;', '')
$extensionText = "section CapacityDistribBTestExtension;`n`n" + $helper + $checkFixture + $joinFixture + $gateFixture + $fixtures + $inputFixtures
if ($extensionText -match 'File\.Contents\s*\(|Excel\.(Workbook|CurrentWorkbook)\s*\(') { throw 'Synthetic fixture package must not contain workbook/file imports.' }
$extensionPath = Join-Path $PSScriptRoot 'CapacityDistribBTests.mez'
$extensionStream = [IO.FileStream]::new($extensionPath, [IO.FileMode]::Create)
$archive = [IO.Compression.ZipArchive]::new($extensionStream, [IO.Compression.ZipArchiveMode]::Create)
try {
    $entry = $archive.CreateEntry('CapacityDistribBTestExtension.pq')
    $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($extensionText) } finally { $writer.Dispose() }
} finally { $archive.Dispose(); $extensionStream.Dispose() }
$queryPath = Join-Path $PSScriptRoot 'fixtures.query.pq'
[IO.File]::WriteAllText($queryPath, $Query, [Text.UTF8Encoding]::new($false))
$pqTest = Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-AIN-B-PQTest-2.155.2/package/tools/PQTest.exe'
if (-not (Test-Path -LiteralPath $pqTest)) { throw 'Microsoft Power Query SDK Tools 2.155.2 are required in the task temporary folder.' }
# The extension contains only the exact helper and synthetic fixtures, never workbook imports.
& $pqTest run-test --extension $extensionPath --queryFile $queryPath --prettyPrint > (Join-Path $PSScriptRoot $ResultName)
$testExitCode = $LASTEXITCODE
$nativeResults = Get-Content -LiteralPath (Join-Path $PSScriptRoot $ResultName) -Raw | ConvertFrom-Json
foreach ($result in $nativeResults) {
    $failures = @($result.Output | Where-Object { $_.Passed -ne $true })
    if ($result.Status -ne 'Passed' -or $failures.Count -gt 0) {
        $testExitCode = 1
        $result | ConvertTo-Json -Depth 15
    } else {
        [ordered]@{ Status = $result.Status; FixtureRows = $result.RowCount; FailedRows = $failures.Count; DataSources = @($result.DataSourceAnalysis).Count; ResultFile = "outputs/ain-b-performance/$ResultName" } | ConvertTo-Json
    }
}
if (@($nativeResults).Count -eq 0) { throw 'PQTest returned no test results.' }
exit $testExitCode
