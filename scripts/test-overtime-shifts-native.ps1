[CmdletBinding()]
param(
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-StaffList-PQTest-2.155.2/package/tools/PQTest.exe'),
    [string] $OutputFolder = 'outputs/overtime-roster-checks/shifts-native-tests',
    [switch] $BuildOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$paths = [ordered]@{
    Baseline = 'outputs/overtime-roster-checks/source-gate/BD-Shifts.before.m'
    Shifts = 'Workflows/ResidentialCare/OvertimeRosterChecks/BD/Shifts.pq'
    Consumer = 'outputs/overtime-roster-checks/source-gate/BD-AllocationByShiftAverage.before.m'
    Fixture = 'Workflows/ResidentialCare/OvertimeRosterChecks/Tests/ShiftsIntegration.pq'
}

# Parse actual shared expressions while respecting quoted text and comments.
function Get-MDefinitions([string] $Text) {
    foreach ($match in [regex]::Matches($Text, '(?m)^shared\s+(?<name>#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=')) {
        $start = $match.Index + $match.Length
        $inString = $false
        $lineComment = $false
        $blockDepth = 0
        $end = -1
        for ($index = $start; $index -lt $Text.Length; $index++) {
            $current = $Text[$index]
            $next = if ($index + 1 -lt $Text.Length) { $Text[$index + 1] } else { [char]0 }
            if ($lineComment) { if ($current -eq "`n") { $lineComment = $false }; continue }
            if ($blockDepth -gt 0) {
                if ($current -eq '/' -and $next -eq '*') { $blockDepth++; $index++ }
                elseif ($current -eq '*' -and $next -eq '/') { $blockDepth--; $index++ }
                continue
            }
            if ($inString) {
                if ($current -eq '"' -and $next -eq '"') { $index++ }
                elseif ($current -eq '"') { $inString = $false }
                continue
            }
            if ($current -eq '/' -and $next -eq '/') { $lineComment = $true; $index++; continue }
            if ($current -eq '/' -and $next -eq '*') { $blockDepth = 1; $index++; continue }
            if ($current -eq '"') { $inString = $true; continue }
            if ($current -eq ';') { $end = $index; break }
        }
        if ($end -lt 0) { throw "Unterminated shared M definition: $($match.Groups['name'].Value)" }
        $token = $match.Groups['name'].Value
        $name = if ($token.StartsWith('#"')) { $token.Substring(2, $token.Length - 3).Replace('""', '"') } else { $token }
        [pscustomobject]@{ Name = $name; Token = $token; Expression = $Text.Substring($start, $end - $start).Trim() }
    }
}

function New-ShiftsEvaluator([object[]] $Definitions, [bool] $WithChecks) {
    $overrides = @{
        'UnitL1PathTABLE' = 'Inputs[PathTable]'
        'IMPORT Table_AllocationExtracted' = 'Inputs[Roster]'
        'IMPORT Intervals' = 'Inputs[IntervalsNavigation]'
        'IMPORT Settings Data' = 'Inputs[SettingsNavigation]'
        'LINK Role Analysis' = 'Inputs[Roles]'
    }
    if ($WithChecks) { $overrides['IMPORT OvertimeRosterChecks'] = 'Inputs[Checks]' }
    foreach ($name in $overrides.Keys) {
        if ($Definitions.Name -notcontains $name) { throw "Missing reviewed Shifts fixture boundary: $name" }
    }
    $bindings = foreach ($definition in $Definitions) {
        $expression = if ($overrides.ContainsKey($definition.Name)) { $overrides[$definition.Name] } else { $definition.Expression }
        $definition.Token + ' = ' + $expression
    }
    $outputs = 'Points = StaffDoubleShifts, Intervals = #"NAMES+INTERVALSLIST"'
    if ($WithChecks) { $outputs += ', Prepared = RosterShiftsWithChecks, Published = OvertimeRosterChecks' }
    '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + "`nin [" + $outputs + ']'
}

function Get-OutputRoot([string] $RelativePath) {
    if ([IO.Path]::IsPathRooted($RelativePath)) { throw 'OutputFolder must be repo-relative.' }
    $resolved = [IO.Path]::GetFullPath((Join-Path $repoRoot $RelativePath))
    $allowed = $repoRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'OutputFolder must remain inside the repository.' }
    $resolved
}

$hashes = [ordered]@{}
$definitions = @{}
foreach ($key in $paths.Keys) {
    $fullPath = Join-Path $repoRoot $paths[$key]
    $hashes[$key] = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash
    $definitions[$key] = @(Get-MDefinitions ([IO.File]::ReadAllText($fullPath)))
    if (-not $definitions[$key].Count -or $definitions[$key].Count -ne @($definitions[$key].Name | Select-Object -Unique).Count) {
        throw "Empty or duplicate shared definitions in $($paths[$key])."
    }
}
$fixture = @($definitions.Fixture | Where-Object Name -eq 'ShiftsIntegrationFixtures')
if ($fixture.Count -ne 1 -or $definitions.Fixture.Count -ne 1) { throw 'Expected one ShiftsIntegrationFixtures function.' }
$consumer = @($definitions.Consumer | Where-Object Name -eq 'ResourceIntervalAllocation')
if ($consumer.Count -ne 1) { throw 'Missing actual ResourceIntervalAllocation definition.' }
$baselineEvaluator = New-ShiftsEvaluator $definitions.Baseline $false
$newEvaluator = New-ShiftsEvaluator $definitions.Shifts $true
$consumerEvaluator = '(Rows as table) as table => let #"EXTRACT NAMES_INTERVALSLIST" = Rows, ResourceIntervalAllocation = ' + $consumer[0].Expression + ' in ResourceIntervalAllocation'
$testSource = "section OvertimeShiftsIntegrationNative;`nshared OvertimeShiftsIntegrationNativeTests = () as table => let`nEvaluateBaseline = " +
    $baselineEvaluator + ",`nEvaluateNew = " + $newEvaluator + ",`nEvaluateConsumer = " + $consumerEvaluator + ",`nFixtures = " +
    $fixture[0].Expression + "`nin Fixtures(EvaluateBaseline, EvaluateNew, EvaluateConsumer);`n"
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage|DataLake|Databricks|Snowflake|PostgreSQL|MySQL|Oracle|SapHana)\.[A-Za-z]+\s*\(|\b(Value\.NativeQuery|Expression\.Evaluate|Extension\.Contents)\s*\(') {
    throw 'Shifts integration fixtures still contain an external connector or dynamic-source call.'
}

$runId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$runFolder = Join-Path (Get-OutputRoot $OutputFolder) $runId
[void][IO.Directory]::CreateDirectory($runFolder)
$encoding = [Text.UTF8Encoding]::new($false)
[IO.File]::WriteAllText((Join-Path $runFolder 'native-fixtures.pq'), $testSource, $encoding)
$summary = [ordered]@{
    Status = 'Built'; Sources = $paths; SHA256 = $hashes; FixtureRows = 0; FailedRows = 0; DataSources = 0
    Scope = 'Actual baseline/new Shifts expressions and existing ResourceIntervalAllocation with synthetic inputs; no workbook or SharePoint access'
    RunFolder = $OutputFolder.TrimEnd('\', '/') + '/' + $runId
}
$temporaryRoot = $null
$runnerProcess = $null
try {
    if ($BuildOnly) { $summary | ConvertTo-Json -Depth 6; return }
    if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) { throw 'Installed Microsoft Power Query SDK Tools PQTest.exe are required.' }
    $probe = [Text.Encoding]::UTF8.GetBytes('ResidentialCare-Overtime-Shifts-PQTest-Probe')
    $encrypted = [Security.Cryptography.ProtectedData]::Protect($probe, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $roundTrip = [Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    if (-not [Linq.Enumerable]::SequenceEqual[byte]($probe, $roundTrip)) { throw 'PQTest profile encryption preflight failed.' }
    $temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('ResidentialCare-Overtime-Shifts-Tests-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($temporaryRoot)
    $extensionPath = Join-Path $temporaryRoot 'OvertimeShiftsTests.mez'
    $queryPath = Join-Path $temporaryRoot 'fixtures.query.pq'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($extensionPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $writer = [IO.StreamWriter]::new($archive.CreateEntry('OvertimeShiftsTests.pq').Open(), $encoding)
        try { $writer.Write($testSource) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
    [IO.File]::WriteAllText($queryPath, 'OvertimeShiftsIntegrationNativeTests()', $encoding)
    $info = [Diagnostics.ProcessStartInfo]::new($PQTestPath)
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = 'Hidden'
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    foreach ($argument in @('-cfp', (Join-Path $temporaryRoot 'synthetic-credentials.bin'), 'run-test', '--extension', $extensionPath, '--queryFile', $queryPath, '--prettyPrint')) {
        $info.ArgumentList.Add($argument)
    }
    $runnerProcess = [Diagnostics.Process]::Start($info)
    $stdout = $runnerProcess.StandardOutput.ReadToEndAsync(); $stderr = $runnerProcess.StandardError.ReadToEndAsync()
    if (-not $runnerProcess.WaitForExit(60000)) { $runnerProcess.Kill($true); throw 'Native Shifts fixtures exceeded 60 seconds.' }
    $nativeOutput = $stdout.GetAwaiter().GetResult()
    [IO.File]::WriteAllText((Join-Path $runFolder 'native-result.json'), $nativeOutput, $encoding)
    [IO.File]::WriteAllText((Join-Path $runFolder 'native-stderr.txt'), $stderr.GetAwaiter().GetResult(), $encoding)
    $jsonStart = $nativeOutput.IndexOf('[')
    if ($jsonStart -lt 0) { throw 'No native JSON results returned.' }
    $results = @($nativeOutput.Substring($jsonStart) | ConvertFrom-Json)
    $rows = @($results | ForEach-Object { @($_.Output) })
    $failedRows = @($rows | Where-Object Passed -ne $true)
    $summary.FixtureRows = $rows.Count; $summary.FailedRows = $failedRows.Count
    $summary.DataSources = @($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
    if ($runnerProcess.ExitCode -ne 0 -or -not $results.Count -or @($results | Where-Object Status -ne 'Passed').Count -or -not $rows.Count -or $failedRows.Count -or $summary.DataSources) {
        $failedRows | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runFolder 'failed-cases.json') -Encoding utf8
        throw 'Native Shifts integration fixtures failed; inspect run artifacts.'
    }
    $summary.Status = 'Passed'
} catch {
    $summary.Status = 'Failed'; $summary['Error'] = $_.Exception.Message
    throw
} finally {
    if ($null -ne $runnerProcess) { if (-not $runnerProcess.HasExited) { $runnerProcess.Kill($true) }; $runnerProcess.Dispose() }
    $unchanged = $true
    foreach ($key in $paths.Keys) {
        if ($hashes[$key] -cne (Get-FileHash -LiteralPath (Join-Path $repoRoot $paths[$key]) -Algorithm SHA256).Hash) { $unchanged = $false }
    }
    $summary['SourcesUnchanged'] = $unchanged
    if (-not $unchanged) { $summary.Status = 'Failed'; $summary['Error'] = 'A source or fixture changed during execution.' }
    $summary | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $runFolder 'summary.json') -Encoding utf8
    if ($null -ne $temporaryRoot -and (Test-Path -LiteralPath $temporaryRoot)) {
        $resolved = [IO.Path]::GetFullPath($temporaryRoot)
        $allowed = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
        if (-not $resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -or -not [IO.Path]::GetFileName($resolved).StartsWith('ResidentialCare-Overtime-Shifts-Tests-')) {
            throw 'Unexpected native test temporary directory; cleanup refused.'
        }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
    if (-not $unchanged) { throw 'Source or fixture changed during native tests.' }
}
$summary | ConvertTo-Json -Depth 7
