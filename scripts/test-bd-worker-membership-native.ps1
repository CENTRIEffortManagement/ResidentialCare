[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $SnapshotJson,
    [Parameter(Mandatory = $true)]
    [string] $BaselineM,
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-AIN-B-PQTest-2.155.2/package/tools/PQTest.exe')
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$sourcePath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m'
$fixturePath = Join-Path $repoRoot 'Workflows/ResidentialCare/Diagnostics/BDWorkerMembership_TEST.pq'

function Get-MDefinitions([string] $text) {
    foreach ($match in [regex]::Matches($text, '(?m)^shared\s+(?<name>#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=')) {
        $start = $match.Index + $match.Length
        $inString = $false
        $lineComment = $false
        $end = -1
        for ($index = $start; $index -lt $text.Length; $index++) {
            $character = $text[$index]
            $next = if ($index + 1 -lt $text.Length) { $text[$index + 1] } else { [char] 0 }
            if ($lineComment) {
                if ($character -eq "`n") { $lineComment = $false }
                continue
            }
            if ($inString) {
                if ($character -eq '"' -and $next -eq '"') { $index++ }
                elseif ($character -eq '"') { $inString = $false }
                continue
            }
            if ($character -eq '/' -and $next -eq '/') { $lineComment = $true; $index++; continue }
            if ($character -eq '"') { $inString = $true; continue }
            if ($character -eq ';') { $end = $index; break }
        }
        if ($end -lt 0) { throw "Unterminated shared definition: $($match.Groups['name'].Value)" }
        $token = $match.Groups['name'].Value
        $name = if ($token.StartsWith('#"')) { $token.Substring(2, $token.Length - 3).Replace('""', '"') } else { $token }
        [pscustomobject] @{ Name = $name; Token = $token; Expression = $text.Substring($start, $end - $start).Trim() }
    }
}

$definitions = @(Get-MDefinitions ([IO.File]::ReadAllText($sourcePath)))
if ($definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) { throw 'Duplicate shared query names.' }
$selectedNames = @('AvailabilityText', 'WorkerRoleKey', 'ReconciledWorkerMembership_Prepare',
    'ReconciledWorkers_Facility_Prepare')
foreach ($name in $selectedNames) {
    if ($definitions.Name -notcontains $name) { throw "Missing membership test target: $name" }
}
$bindings = @('#"IMPORT Reconciled Workers" = Inputs[Navigation]', 'Unit = "BD"',
    '#"IMPORT Availability Leave Source" = #table({"Payroll Code", "Facility-Abbrev", "Employee Name"}, {})')
$bindings += @($definitions | Where-Object Name -in $selectedNames | ForEach-Object { $_.Token + ' = ' + $_.Expression })
# Evaluate the actual membership grouping steps before unrelated role, name, contract and interval inputs.
$workerPreparation = ($definitions | Where-Object Name -eq 'ReconciledWorkers_Prepare').Expression
$cut = $workerPreparation.IndexOf('    // Join the worker')
if ($cut -lt 0) { throw 'The membership grouping boundary was not found.' }
$membershipPrefix = $workerPreparation.Substring(0, $cut).TrimEnd()
if (-not $membershipPrefix.EndsWith(',')) { throw 'Expected a comma before the role-mapping boundary.' }
$membershipExpression = $membershipPrefix.Substring(0, $membershipPrefix.Length - 1) + "`nin RemovedSourceRoleLists"
$bindings += 'ReconciledWorkers_Prepare = ' + $membershipExpression
$evaluate = '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + @'

in [Prepared = ReconciledWorkerMembership_Prepare, Grouped = ReconciledWorkers_Prepare,
    Facility = ReconciledWorkers_Facility_Prepare]
'@

# Exercise the old reader against the identical navigation snapshot to reproduce the original error.
$baselineDefinitions = @(Get-MDefinitions ([IO.File]::ReadAllText((Resolve-Path -LiteralPath $BaselineM).Path)))
$legacy = @($baselineDefinitions | Where-Object Name -eq 'IMPORT Reconciled Workers')
if ($legacy.Count -ne 1) { throw 'Expected one baseline reconciliation reader.' }
$legacyExpression = $legacy[0].Expression
$fileRead = 'File.Contents(SourcePath)'
$workbookRead = 'Excel.Workbook(SourceBinary, null, true)'
if (-not $legacyExpression.Contains($fileRead) -or -not $legacyExpression.Contains($workbookRead)) {
    throw 'Baseline connector expressions do not match the approved fixture injection points.'
}
$legacyExpression = $legacyExpression.Replace($fileRead, 'null').Replace($workbookRead, 'InputNavigation')
$legacyEvaluator = '(InputNavigation as table) as table => let #"FilePath - 1Input" = ".", AvailabilityText = ' +
    ($definitions | Where-Object Name -eq 'AvailabilityText').Expression + ', Legacy = ' + $legacyExpression + ' in Legacy'

$workerChecks = ($definitions | Where-Object Name -eq 'ReconciledWorkers_CHECK').Expression
$validationStart = $workerChecks.IndexOf('    InvalidMembershipRows = ')
$validationEnd = $workerChecks.IndexOf('    DuplicateCount = ')
if ($validationStart -lt 0 -or $validationEnd -le $validationStart) { throw 'Membership validation boundary not found.' }
$validationBinding = $workerChecks.Substring($validationStart, $validationEnd - $validationStart).TrimEnd()
if (-not $validationBinding.EndsWith(',')) { throw 'Expected a comma after membership validation.' }
$invalidMembershipEvaluator = '(Workers as table) as table => let ' +
    $validationBinding.Substring(0, $validationBinding.Length - 1) + ' in InvalidMembershipRows'

$snapshotPath = (Resolve-Path -LiteralPath $SnapshotJson).Path
$snapshot = @(Get-Content -LiteralPath $snapshotPath -Raw | ConvertFrom-Json)
if ($snapshot.Count -eq 0) { throw 'The captured reconciliation snapshot is empty.' }
$snapshotColumns = @('EmployeeID', 'Facility-Abbrev', 'Employee Roster Name', 'EmploymentType', 'Role',
    'In Roster', 'In Availabilities', 'In Leave Balances', 'In Records')
$snapshotJsonText = ConvertTo-Json -InputObject @($snapshot | Select-Object $snapshotColumns) -Depth 6 -Compress
$encodedSnapshot = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($snapshotJsonText))
$snapshotExpression = 'Table.FromRecords(Json.Document(Binary.FromText("' + $encodedSnapshot + '", BinaryEncoding.Base64)))'
$fixture = @(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)) | Where-Object Name -eq 'BDWorkerMembershipFixtures')
if ($fixture.Count -ne 1) { throw 'Expected one membership fixture function.' }
$testSource = "section BDWorkerMembershipNative;`nshared BDWorkerMembershipTests = () as table => let`nEvaluate = " +
    $evaluate + ",`nEvaluateLegacy = " + $legacyEvaluator + ",`nInvalidMembershipRows = " + $invalidMembershipEvaluator +
    ",`nCapturedWorkers = " + $snapshotExpression + ",`nFixtures = " + $fixture[0].Expression +
    "`nin Fixtures(Evaluate, EvaluateLegacy, InvalidMembershipRows, CapturedWorkers);`n"
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(|\bValue\.NativeQuery\s*\(') {
    throw 'Membership fixtures must not contain workbook, file or network data access.'
}
if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) { throw 'The installed Microsoft Power Query SDK Tools are required.' }

$outputFolder = Join-Path $repoRoot ('outputs/bd-worker-membership-change1/' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = New-Item -ItemType Directory -Path $outputFolder
$extensionPath = Join-Path $outputFolder 'BDWorkerMembershipTests.mez'
$queryPath = Join-Path $outputFolder 'fixtures.query.pq'
$resultPath = Join-Path $outputFolder 'native-result.json'
$errorPath = Join-Path $outputFolder 'native-stderr.txt'
$credentialsPath = Join-Path $outputFolder 'fixture-credentials.bin'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::Open($extensionPath, [IO.Compression.ZipArchiveMode]::Create)
try {
    $entry = $archive.CreateEntry('BDWorkerMembershipTests.pq')
    $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($testSource) } finally { $writer.Dispose() }
} finally { $archive.Dispose() }
[IO.File]::WriteAllText($queryPath, 'BDWorkerMembershipTests()', [Text.UTF8Encoding]::new($false))

# The engine initializes Windows profile encryption even for data-source-free fixtures.
$probe = [Text.Encoding]::UTF8.GetBytes('BD-worker-membership-native-test')
$encrypted = [Security.Cryptography.ProtectedData]::Protect($probe, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
$roundTrip = [Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
if (-not [Linq.Enumerable]::SequenceEqual[byte]($probe, $roundTrip)) { throw 'Power Query test engine profile probe failed.' }

$process = $null
try {
    $arguments = '-cfp "' + $credentialsPath + '" run-test --extension "' + $extensionPath +
        '" --queryFile "' + $queryPath + '" --prettyPrint'
    $process = Start-Process -FilePath $PQTestPath -ArgumentList $arguments -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput $resultPath -RedirectStandardError $errorPath
    if (-not $process.WaitForExit(60000)) {
        $process.Kill($true)
        throw 'The isolated membership test exceeded 60 seconds.'
    }
    $lines = @(Get-Content -LiteralPath $resultPath)
    $jsonStart = -1
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if ($lines[$index].TrimStart().StartsWith('[') -or $lines[$index].TrimStart().StartsWith('{')) { $jsonStart = $index; break }
    }
    if ($jsonStart -lt 0) { throw "No native result was returned. Inspect $resultPath and $errorPath." }
    $results = @(($lines[$jsonStart..($lines.Count - 1)] -join "`n") | ConvertFrom-Json)
    $failedQueries = @($results | Where-Object Status -ne 'Passed')
    $failedRows = @($results | ForEach-Object { @($_.Output | Where-Object Passed -ne $true) })
    $dataSources = @($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
    $summary = [ordered] @{
        Status = if ($process.ExitCode -eq 0 -and $results.Count -gt 0 -and $failedQueries.Count -eq 0 -and $failedRows.Count -eq 0 -and $dataSources -eq 0) { 'Passed' } else { 'Failed' }
        FixtureRows = [int] ($results | Measure-Object RowCount -Sum).Sum
        FailedRows = $failedRows.Count
        CapturedWorkerRows = $snapshot.Count
        DataSources = $dataSources
        Source = [IO.Path]::GetRelativePath($repoRoot, $sourcePath).Replace('\', '/')
        SourceSHA256 = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
        SnapshotSHA256 = (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash
        OutputFolder = [IO.Path]::GetRelativePath($repoRoot, $outputFolder).Replace('\', '/')
        Scope = 'Production membership adapter and grouping steps; captured reconciliation rows and focused boundary cases; no Excel refresh or external access'
    }
    $summary | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $outputFolder 'summary.json') -Encoding UTF8
    $summary | ConvertTo-Json
    if ($summary.Status -ne 'Passed') {
        $results | ConvertTo-Json -Depth 12
        throw 'The isolated membership test failed.'
    }
}
finally {
    if ($null -ne $process) {
        if (-not $process.HasExited) { $process.Kill($true) }
        $process.Dispose()
    }
}
