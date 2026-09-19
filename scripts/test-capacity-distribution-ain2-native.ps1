[CmdletBinding()]
param(
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-AIN-B-PQTest-2.155.2/package/tools/PQTest.exe'),
    [string] $Query = 'CapacityDistribution_AIN2_CHECK'
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$fixturePath = Join-Path $repoRoot 'Workflows/ResidentialCare/Diagnostics/CapacityDistribution_AIN2_TEST.pq'
$existingPQTestIds = @(Get-Process -Name 'PQTest' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$previousErrorMode = $null

# Suppress Windows crash-dialog UI from a failing test executable; failures must return to this
# non-interactive wrapper as output/exit status instead of blocking the caller.
if (-not ('CapacityDistributionNativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CapacityDistributionNativeMethods {
    [DllImport("kernel32.dll")]
    public static extern uint SetErrorMode(uint uMode);
}
'@
}
$SEM_NOGPFAULTERRORBOX = [uint32]0x0002
$previousErrorMode = [CapacityDistributionNativeMethods]::SetErrorMode($SEM_NOGPFAULTERRORBOX)

if (-not (Test-Path -LiteralPath $fixturePath -PathType Leaf)) {
    throw "Missing native fixture source: $fixturePath"
}
if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) {
    throw "Microsoft Power Query SDK Tools 2.155.2 were not found at: $PQTestPath"
}

# PQTest initializes a CurrentUser DPAPI credential store even for this data-source-free suite.
# Fail before launching it when the host has no usable user profile (for example, a restricted
# automation token), because that host condition otherwise leaves a crashing PQTest process.
try {
    $probeBytes = [Text.Encoding]::UTF8.GetBytes('ResidentialCare-AIN2-PQTest-DPAPI-Probe')
    $protectedProbe = [Security.Cryptography.ProtectedData]::Protect(
        $probeBytes,
        $null,
        [Security.Cryptography.DataProtectionScope]::CurrentUser
    )
    $roundTripProbe = [Security.Cryptography.ProtectedData]::Unprotect(
        $protectedProbe,
        $null,
        [Security.Cryptography.DataProtectionScope]::CurrentUser
    )
    if (-not [Linq.Enumerable]::SequenceEqual[byte]($probeBytes, $roundTripProbe)) {
        throw 'The DPAPI probe did not round-trip.'
    }
}
catch {
    throw "PQTest requires a loaded Windows user profile for its disposable credential store. Run this script in the signed-in user context. $($_.Exception.Message)"
}

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('ResidentialCare-CapacityDistribution-AIN2-Tests-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temporaryRoot

try {
    $extensionPath = Join-Path $temporaryRoot 'CapacityDistributionAIN2Tests.mez'
    $queryPath = Join-Path $temporaryRoot 'fixtures.query.pq'
    $resultPath = Join-Path $temporaryRoot 'result.json'
    $credentialsPath = Join-Path $temporaryRoot 'credentials.bin'
    $fixtureSource = [IO.File]::ReadAllText($fixturePath)

    # PQTest's discovery pass can force exported checks that deliberately throw.
    # When the caller requests the diagnostic table, omit only the final fail-closed
    # wrapper so its individual result rows remain inspectable.
    if ($Query -eq 'CapacityDistribution_AIN2_TEST') {
        $checkMarker = '// Query: CapacityDistribution_AIN2_CHECK'
        $checkStart = $fixtureSource.IndexOf($checkMarker, [StringComparison]::Ordinal)
        if ($checkStart -lt 0) { throw "Cannot locate $checkMarker in the fixture source." }
        $fixtureSource = $fixtureSource.Substring(0, $checkStart)
    }

    if ($fixtureSource -match 'File\.Contents\s*\(|Excel\.(Workbook|CurrentWorkbook)\s*\(') {
        throw 'AIN2 native fixtures must remain synthetic and must not contain workbook or file imports.'
    }

    $extensionStream = [IO.FileStream]::new($extensionPath, [IO.FileMode]::CreateNew)
    $archive = [IO.Compression.ZipArchive]::new($extensionStream, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $entry = $archive.CreateEntry('CapacityDistributionAIN2Tests.pq')
        $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
        try {
            $writer.Write($fixtureSource)
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $archive.Dispose()
        $extensionStream.Dispose()
    }

    [IO.File]::WriteAllText($queryPath, $Query, [Text.UTF8Encoding]::new($false))
    # Use a disposable credential-cache path so this data-source-free suite
    # neither reads nor writes the user's persistent PQTest credential store.
    & $PQTestPath -cfp $credentialsPath run-test --extension $extensionPath --queryFile $queryPath --prettyPrint > $resultPath
    $testExitCode = $LASTEXITCODE
    $resultLines = @(Get-Content -LiteralPath $resultPath)
    $jsonStart = -1
    for ($lineIndex = 0; $lineIndex -lt $resultLines.Count; $lineIndex++) {
        $trimmed = $resultLines[$lineIndex].TrimStart()
        if ($trimmed.StartsWith('[') -or $trimmed.StartsWith('{')) { $jsonStart = $lineIndex; break }
    }
    if ($jsonStart -lt 0) {
        throw "PQTest returned no JSON result: $($resultLines -join ' ')"
    }
    $results = @(($resultLines[$jsonStart..($resultLines.Count - 1)] -join "`n") | ConvertFrom-Json)

    if ($results.Count -eq 0) {
        throw 'PQTest returned no AIN2 fixture result.'
    }

    $failedRows = @($results | ForEach-Object { @($_.Output | Where-Object { $_.Passed -ne $true }) })
    $failedExecutions = @($results | Where-Object { $_.Status -ne 'Passed' })
    if ($testExitCode -ne 0 -or $failedExecutions.Count -gt 0 -or $failedRows.Count -gt 0) {
        $results | ConvertTo-Json -Depth 15
        exit 1
    }

    [ordered]@{
        Status = 'Passed'
        FixtureRows = [int]($results | Measure-Object -Property RowCount -Sum).Sum
        FailedRows = 0
        DataSources = @($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
        Scope = 'Synthetic Power Query fixtures only; no workbook inspection or refresh'
    } | ConvertTo-Json
}
finally {
    # PQTest has occasionally returned its JSON before a child process completed. Remove only
    # processes created by this invocation so native tests cannot leave workers or crash dialogs.
    Start-Sleep -Milliseconds 250
    $newPQTestProcesses = @(Get-Process -Name 'PQTest' -ErrorAction SilentlyContinue | Where-Object {
        $existingPQTestIds -notcontains $_.Id
    })
    foreach ($process in $newPQTestProcesses) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 100
    $leakedPQTestProcesses = @(Get-Process -Name 'PQTest' -ErrorAction SilentlyContinue | Where-Object {
        $existingPQTestIds -notcontains $_.Id
    })
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
    if ($null -ne $previousErrorMode) {
        $null = [CapacityDistributionNativeMethods]::SetErrorMode([uint32]$previousErrorMode)
    }
    if ($leakedPQTestProcesses.Count -gt 0) {
        throw "PQTest left process IDs running after cleanup: $($leakedPQTestProcesses.Id -join ', ')."
    }
}
