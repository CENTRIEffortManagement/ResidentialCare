[CmdletBinding()]
param(
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-StaffList-PQTest-2.155.2/package/tools/PQTest.exe'),
    [string] $EnvironmentConfiguration = '',
    [string] $OutputFolder = 'outputs/overtime-roster-checks/native-tests',
    [switch] $BuildOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
$sourceRelativePath = 'Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq'
$fixtureRelativePath = 'Workflows/ResidentialCare/OvertimeRosterChecks/Tests/OvertimeRosterChecks_TEST.pq'
$sourcePath = Join-Path $repoRoot $sourceRelativePath
$fixturePath = Join-Path $repoRoot $fixtureRelativePath

# Read actual shared definitions without executing any connector or touching a workbook.
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
            if ($lineComment) {
                if ($current -eq "`n") { $lineComment = $false }
                continue
            }
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

function Assert-NoExternalSources([string] $Text) {
    # Reject every connector still present after the explicit fixture substitutions.
    # Data-source analysis is checked independently after native execution.
    if ($Text -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage|DataLake|Databricks|Snowflake|PostgreSQL|MySQL|Oracle|SapHana)\.[A-Za-z]+\s*\(|\b(Value\.NativeQuery|Expression\.Evaluate|Extension\.Contents)\s*\(') {
        throw 'Native overtime fixtures contain an unapproved file, workbook, network or dynamic-source call.'
    }
}

function Get-RepoOutputPath([string] $RelativePath) {
    if ([IO.Path]::IsPathRooted($RelativePath)) { throw 'OutputFolder must be repo-relative.' }
    $resolved = [IO.Path]::GetFullPath((Join-Path $repoRoot $RelativePath))
    $allowedRoot = $repoRoot.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'OutputFolder must remain inside the repository.'
    }
    $resolved
}

function New-OvertimeNativeEvaluator([object[]] $Definitions) {
    # Only these reviewed landing boundaries may be replaced. Every calculation
    # expression remains the actual production source rather than a test rewrite.
    $overrides = @{
        'INPUT OvertimeRosterChecks FilePathUrl' = 'Inputs[PathInput]'
        'IMPORT OvertimeRosterChecks CentriSyncPaths' = 'Inputs[PathMapping]'
        'IMPORT Published Roster' = '(unit as text) as table => if unit = "BD" then Inputs[RosterBD] else if unit = "TE" then Inputs[RosterTE] else if unit = "JH-RY" then Inputs[RosterJHRY] else error "Unapproved fixture unit."'
        'IMPORT Workers' = '(unit as text) as table => if unit = "BD" then Inputs[WorkersBD] else if unit = "TE" then Inputs[WorkersTE] else if unit = "JH-RY" then Inputs[WorkersJHRY] else error "Unapproved fixture unit."'
        'IMPORT Settings Data' = '(unit as text) as table => if unit = "BD" then Inputs[SettingsBD] else if unit = "TE" then Inputs[SettingsTE] else if unit = "JH-RY" then Inputs[SettingsJHRY] else error "Unapproved fixture unit."'
        'IMPORT OvertimeRosterChecks Whiddon Lists' = 'error "SharePoint navigation is disabled in native fixtures."'
        'LINK OvertimeRosterThreshold' = 'Inputs[Thresholds]'
        'LINK OvertimeRosterChecks Roles' = 'Inputs[Roles]'
        'LINK OvertimeRosterChecks Facilities' = 'Inputs[Facilities]'
    }
    foreach ($name in $overrides.Keys) {
        if ($Definitions.Name -notcontains $name) { throw "Missing reviewed native fixture landing: $name" }
    }
    $bindings = foreach ($definition in $Definitions) {
        $expression = if ($overrides.ContainsKey($definition.Name)) { $overrides[$definition.Name] } else { $definition.Expression }
        $definition.Token + ' = ' + $expression
    }
    '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + @'

in [Path = OvertimeRosterChecks_Path,
    Anchor = OvertimeRosterChecks_FortnightStartDate,
    Shifts = OvertimeRosterChecks_ShiftResults,
    Fortnights = OvertimeRosterChecks_FortnightResults,
    Segments = OvertimeRosterChecks_ExcessSegments]
'@
}

$sourceHashes = [ordered]@{
    Source = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
    Fixture = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash
}
$definitions = @(Get-MDefinitions ([IO.File]::ReadAllText($sourcePath)))
if (-not $definitions.Count -or $definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) {
    throw 'Production source contains no shared definitions or duplicate shared names.'
}
$fixtureDefinitions = @(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)))
$fixture = @($fixtureDefinitions | Where-Object Name -eq 'OvertimeRosterChecksFixtures')
if ($fixture.Count -ne 1 -or $fixtureDefinitions.Count -ne 1) { throw 'Expected one OvertimeRosterChecksFixtures shared function.' }
$evaluator = New-OvertimeNativeEvaluator $definitions
$testSource = "section OvertimeRosterChecksNative;`nshared OvertimeRosterChecksNativeTests = () as table => let`nEvaluate = " +
    $evaluator + ",`nFixtures = " + $fixture[0].Expression + "`nin Fixtures(Evaluate);`n"
Assert-NoExternalSources $testSource

$outputRoot = Get-RepoOutputPath $OutputFolder
$runId = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$runFolder = Join-Path $outputRoot $runId
[void][IO.Directory]::CreateDirectory($runFolder)
$encoding = [Text.UTF8Encoding]::new($false)
[IO.File]::WriteAllText((Join-Path $runFolder 'native-fixtures.pq'), $testSource, $encoding)
$summary = [ordered]@{
    Status = 'Built'
    Source = $sourceRelativePath
    Fixture = $fixtureRelativePath
    SourceSHA256 = $sourceHashes.Source
    FixtureSHA256 = $sourceHashes.Fixture
    FixtureRows = 0
    FailedRows = 0
    DataSources = 0
    TimeoutSeconds = 60
    Scope = 'Actual new OvertimeRosterChecks M definitions with synthetic landed inputs only; no workbook or SharePoint access'
    RunFolder = $OutputFolder.TrimEnd('\', '/') + '/' + $runId
}
$temporaryRoot = $null
$runnerProcess = $null
try {
    if ($BuildOnly) {
        $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $runFolder 'summary.json') -Encoding utf8
        $summary | ConvertTo-Json -Depth 6
        return
    }
    if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) {
        throw 'Installed Microsoft Power Query SDK Tools PQTest.exe are required.'
    }
    # The native engine needs the signed-in Windows encryption profile even
    # though the generated fixtures have no live data sources or credentials.
    $probe = [Text.Encoding]::UTF8.GetBytes('ResidentialCare-OvertimeRosterChecks-PQTest-Probe')
    $encrypted = [Security.Cryptography.ProtectedData]::Protect($probe, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $roundTrip = [Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
    if (-not [Linq.Enumerable]::SequenceEqual[byte]($probe, $roundTrip)) { throw 'PQTest profile encryption preflight failed.' }

    $temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('ResidentialCare-OvertimeRosterChecks-Tests-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($temporaryRoot)
    $extensionPath = Join-Path $temporaryRoot 'OvertimeRosterChecksTests.mez'
    $queryPath = Join-Path $temporaryRoot 'fixtures.query.pq'
    $credentialsPath = Join-Path $temporaryRoot 'synthetic-credentials.bin'
    $resultPath = Join-Path $runFolder 'native-result.json'
    $errorPath = Join-Path $runFolder 'native-stderr.txt'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($extensionPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entry = $archive.CreateEntry('OvertimeRosterChecksTests.pq')
        $writer = [IO.StreamWriter]::new($entry.Open(), $encoding)
        try { $writer.Write($testSource) } finally { $writer.Dispose() }
    } finally { $archive.Dispose() }
    [IO.File]::WriteAllText($queryPath, 'OvertimeRosterChecksNativeTests()', $encoding)
    $info = [Diagnostics.ProcessStartInfo]::new($PQTestPath)
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.WindowStyle = 'Hidden'
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in @('-cfp', $credentialsPath, 'run-test', '--extension', $extensionPath, '--queryFile', $queryPath, '--prettyPrint')) {
        $info.ArgumentList.Add($argument)
    }
    if (-not [string]::IsNullOrWhiteSpace($EnvironmentConfiguration)) {
        if (-not (Test-Path -LiteralPath $EnvironmentConfiguration -PathType Leaf)) { throw 'EnvironmentConfiguration file does not exist.' }
        $info.ArgumentList.Add('--environmentConfiguration')
        $info.ArgumentList.Add([IO.Path]::GetFullPath($EnvironmentConfiguration))
    }
    $runnerProcess = [Diagnostics.Process]::Start($info)
    $stdout = $runnerProcess.StandardOutput.ReadToEndAsync()
    $stderr = $runnerProcess.StandardError.ReadToEndAsync()
    if (-not $runnerProcess.WaitForExit(60000)) {
        $runnerProcess.Kill($true)
        throw 'Isolated native overtime fixtures exceeded 60 seconds.'
    }
    $nativeOutput = $stdout.GetAwaiter().GetResult()
    [IO.File]::WriteAllText($resultPath, $nativeOutput, $encoding)
    [IO.File]::WriteAllText($errorPath, $stderr.GetAwaiter().GetResult(), $encoding)
    $jsonStart = $nativeOutput.IndexOf('[')
    if ($jsonStart -lt 0) { throw 'No native JSON results returned; inspect native-result.json and native-stderr.txt.' }
    $results = @($nativeOutput.Substring($jsonStart) | ConvertFrom-Json)
    $queryFailures = @($results | Where-Object Status -ne 'Passed')
    $passedQueries = @($results | Where-Object Status -eq 'Passed')
    $rows = @($passedQueries | ForEach-Object { @($_.Output) })
    $failedRows = @($rows | Where-Object Passed -ne $true)
    $dataSources = @($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
    $summary.FixtureRows = $rows.Count
    $summary.FailedRows = $failedRows.Count
    $summary.DataSources = $dataSources
    if ($runnerProcess.ExitCode -ne 0 -or -not $results.Count -or $queryFailures.Count -or -not $rows.Count -or $failedRows.Count -or $dataSources) {
        $failedRows | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runFolder 'failed-cases.json') -Encoding utf8
        throw 'Native overtime roster fixtures failed; inspect run artifacts.'
    }
    $summary.Status = 'Passed'
} catch {
    $summary.Status = 'Failed'
    $summary['Error'] = $_.Exception.Message
    throw
} finally {
    if ($null -ne $runnerProcess) {
        if (-not $runnerProcess.HasExited) { $runnerProcess.Kill($true) }
        $runnerProcess.Dispose()
    }
    $afterHashes = @{
        Source = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
        Fixture = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash
    }
    $unchanged = $sourceHashes.Source -ceq $afterHashes.Source -and $sourceHashes.Fixture -ceq $afterHashes.Fixture
    $summary['SourcesUnchanged'] = $unchanged
    if (-not $unchanged) { $summary.Status = 'Failed'; $summary['Error'] = 'Production source or fixture changed during execution.' }
    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $runFolder 'summary.json') -Encoding utf8
    if ($null -ne $temporaryRoot -and (Test-Path -LiteralPath $temporaryRoot)) {
        # Verify this exact invocation's temporary directory before recursive cleanup.
        $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
        $allowedParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
        if (-not $resolvedTemporaryRoot.StartsWith($allowedParent, [StringComparison]::OrdinalIgnoreCase) -or
            -not [IO.Path]::GetFileName($resolvedTemporaryRoot).StartsWith('ResidentialCare-OvertimeRosterChecks-Tests-')) {
            throw 'Unexpected native-test temporary directory; cleanup refused.'
        }
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
    }
    if (-not $unchanged) { throw 'Production source or fixture changed during native tests.' }
}
$summary | ConvertTo-Json -Depth 6
