[CmdletBinding()]
param(
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-AIN-B-PQTest-2.155.2/package/tools/PQTest.exe')
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$availabilityPath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m'
$staffListPath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/StaffListMaster.xlsx_PowerQuery.m'
$fixturePath = Join-Path $repoRoot 'Workflows/ResidentialCare/Diagnostics/BDRosterShiftCaps_TEST.pq'

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

# Select the real changed calculations and inject synthetic inputs at their existing dependency boundaries.
# Unselected connector queries never enter the generated extension; no production data is accessed.
function New-NativeEvaluator {
    param([string] $Path, [string[]] $Names, [string[]] $InputBindings, [string] $Output)
    $definitions = @(Get-MDefinitions ([IO.File]::ReadAllText($Path)))
    if ($definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) { throw "Duplicate query names in $Path" }
    foreach ($name in $Names) {
        if ($definitions.Name -notcontains $name) { throw "Missing roster-cap fixture target: $name" }
    }
    $bindings = @($InputBindings) + @($definitions | Where-Object Name -in $Names | ForEach-Object {
        $_.Token + ' = ' + $_.Expression
    })
    '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + "`nin " + $Output
}

$availabilityNames = @('AvailabilityText', 'CasualRosterShiftCapsByEmployee', 'Availability-StaffList',
    'CasualRosterShiftCapExceptions', 'RoleResDayAvailabilityCapped', 'RoleAvailabilityCapped')
$availabilityInputs = @(
    '#"IMPORT CasualRosterShiftCaps" = Inputs[Caps]',
    'ReconciledWorkers_Eligible = Inputs[Staff]',
    'ReconciledWorkers_CHECK = #table({"Check", "Status"}, {{"Synthetic identity inputs", "Pass"}})',
    '#"EXTRACT MaxAvailability" = Inputs[Maximum]',
    'ResDayShift = Inputs[DayShifts]',
    'AllocationRoleShares = Inputs[Shares]',
    'AllocationRoleShares_CHECK = #table({"Check", "Status"}, {{"Synthetic role shares", "Pass"}})'
)
$availabilityOutput = '[Staff = #"Availability-StaffList", Lookup = CasualRosterShiftCapsByEmployee, ' +
    'Exceptions = CasualRosterShiftCapExceptions, RoleResources = RoleResDayAvailabilityCapped, Roles = RoleAvailabilityCapped]'
$staffListNames = @('StaffListText', 'StaffListCheckResult', 'EXTRACT Availability Staff List',
    'ResourceContract', 'ResourceContract_CHECK', 'Masterlist')
$staffListInputs = @(
    'Masterlist_Base = Inputs[Staff]',
    '#"IMPORT Capacity Shift Availability" = #table({"Item", "Kind", "Data"}, {{"Availability-StaffList", "Sheet", Table.DemoteHeaders(Inputs[Staff])}})',
    '#"EXTRACT StaffList RosterDates" = #table(type table [Date = date], List.Transform({0..27}, each {Date.AddDays(#date(2026, 10, 1), _)}))',
    '#"EXTRACT StaffList ShiftDuration" = 7.6',
    '#"EXTRACT StaffList MaxAvailability" = Inputs[Maximum]'
)
$staffListOutput = '[Contracts = ResourceContract, Checks = ResourceContract_CHECK, Master = Masterlist]'
$availabilityEvaluator = New-NativeEvaluator $availabilityPath $availabilityNames $availabilityInputs $availabilityOutput
$staffListEvaluator = New-NativeEvaluator $staffListPath $staffListNames $staffListInputs $staffListOutput
$fixture = @(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)) | Where-Object Name -eq 'BDRosterShiftCapsFixtures')
if ($fixture.Count -ne 1) { throw 'Expected one BDRosterShiftCapsFixtures function.' }
$testSource = "section BDRosterShiftCapsNative;`nshared BDRosterShiftCapsTests = () as table => let`nEvaluateAvailability = " +
    $availabilityEvaluator + ",`nEvaluateStaffList = " + $staffListEvaluator + ",`nFixtures = " + $fixture[0].Expression +
    "`nin Fixtures(EvaluateAvailability, EvaluateStaffList);`n"
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(|\bValue\.NativeQuery\s*\(') {
    throw 'Native roster-cap tests must not contain file, workbook or network data access.'
}
if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) { throw 'Microsoft Power Query SDK Tools 2.155.2 are required.' }

# PQTest initializes Windows profile encryption even for synthetic, data-source-free fixtures.
$probe = [Text.Encoding]::UTF8.GetBytes('ResidentialCare-BDRosterShiftCaps-PQTest-Probe')
$encrypted = [Security.Cryptography.ProtectedData]::Protect($probe, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
$roundTrip = [Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
if (-not [Linq.Enumerable]::SequenceEqual[byte]($probe, $roundTrip)) { throw 'PQTest profile probe failed.' }

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('ResidentialCare-BDRosterShiftCaps-Tests-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temporaryRoot
$runnerProcess = $null
try {
    $extensionPath = Join-Path $temporaryRoot 'BDRosterShiftCapsTests.mez'
    $queryPath = Join-Path $temporaryRoot 'fixtures.query.pq'
    $resultPath = Join-Path $temporaryRoot 'result.json'
    $errorPath = Join-Path $temporaryRoot 'stderr.txt'
    $credentialsPath = Join-Path $temporaryRoot 'credentials.bin'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($extensionPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entry = $archive.CreateEntry('BDRosterShiftCapsTests.pq')
        $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write($testSource) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
    [IO.File]::WriteAllText($queryPath, 'BDRosterShiftCapsTests()', [Text.UTF8Encoding]::new($false))
    $arguments = '-cfp "' + $credentialsPath + '" run-test --extension "' + $extensionPath +
        '" --queryFile "' + $queryPath + '" --prettyPrint'
    $runnerProcess = Start-Process -FilePath $PQTestPath -ArgumentList $arguments -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput $resultPath -RedirectStandardError $errorPath
    if (-not $runnerProcess.WaitForExit(60000)) {
        $runnerProcess.Kill($true)
        throw 'Power Query roster-cap fixture execution exceeded 60 seconds.'
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
    $dataSources = @($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
    if ($runnerProcess.ExitCode -ne 0 -or $results.Count -eq 0 -or $failures.Count -gt 0 -or $failedRows.Count -gt 0 -or $dataSources -gt 0) {
        $results | ConvertTo-Json -Depth 15
        throw 'BD roster shift-cap native fixtures failed.'
    }
    [ordered]@{
        Status = 'Passed'
        FixtureRows = [int]($results | Measure-Object -Property RowCount -Sum).Sum
        FailedRows = 0
        DataSources = $dataSources
        Scope = 'Actual BD Availability and StaffListMaster roster-cap calculations with synthetic inputs; no external access'
    } | ConvertTo-Json
}
finally {
    if ($null -ne $runnerProcess) {
        if (-not $runnerProcess.HasExited) { $runnerProcess.Kill($true) }
        $runnerProcess.Dispose()
    }
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $allowedParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolvedTemporaryRoot.StartsWith($allowedParent, [StringComparison]::OrdinalIgnoreCase) -or
        -not ([IO.Path]::GetFileName($resolvedTemporaryRoot)).StartsWith('ResidentialCare-BDRosterShiftCaps-Tests-')) {
        throw 'Unexpected test temporary directory; cleanup refused.'
    }
    Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
}
