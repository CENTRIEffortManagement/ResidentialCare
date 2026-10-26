[CmdletBinding()]
param(
    [string] $SourcePath = 'CLIENT/DATExx-Whiddon/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m',
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-StaffList-PQTest-2.155.2/package/tools/PQTest.exe'),
    [switch] $PrepareOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))

function Get-MDefinitions([string] $Text) {
    foreach ($match in [regex]::Matches($Text, '(?m)^shared\s+(?<name>#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=')) {
        $start = $match.Index + $match.Length
        $inString = $false
        $lineComment = $false
        $blockComment = $false
        $end = -1
        for ($index = $start; $index -lt $Text.Length; $index++) {
            $character = $Text[$index]
            $next = if ($index + 1 -lt $Text.Length) { $Text[$index + 1] } else { [char]0 }
            if ($lineComment) { if ($character -eq "`n") { $lineComment = $false }; continue }
            if ($blockComment) { if ($character -eq '*' -and $next -eq '/') { $blockComment = $false; $index++ }; continue }
            if ($inString) {
                if ($character -eq '"' -and $next -eq '"') { $index++ }
                elseif ($character -eq '"') { $inString = $false }
                continue
            }
            if ($character -eq '/' -and $next -eq '/') { $lineComment = $true; $index++; continue }
            if ($character -eq '/' -and $next -eq '*') { $blockComment = $true; $index++; continue }
            if ($character -eq '"') { $inString = $true; continue }
            if ($character -eq ';') { $end = $index; break }
        }
        if ($end -lt 0) { throw "Unterminated M definition: $($match.Groups['name'].Value)" }
        $token = $match.Groups['name'].Value
        $name = if ($token.StartsWith('#"')) { $token.Substring(2, $token.Length - 3).Replace('""', '"') } else { $token }
        [pscustomobject]@{Name=$name; Token=$token; Expression=$Text.Substring($start,$end-$start).Trim()}
    }
}

function Get-MReferences([string] $Expression) {
    $pattern = '(?s)//[^\r\n]*|/\*.*?\*/|#"(?:[^"]|"")*"|"(?:[^"]|"")*"|(?<![\w.])[A-Za-z_][A-Za-z0-9_]*(?![\w.])'
    foreach ($match in [regex]::Matches($Expression, $pattern)) {
        $token = $match.Value
        if ($token.StartsWith('#"')) { $token.Substring(2, $token.Length - 3).Replace('""', '"') }
        elseif ($token -match '^[A-Za-z_]') { $token }
    }
}

function Get-Bindings($Definitions, $Overrides, [string[]] $Targets) {
    $map = @{}
    foreach ($definition in $Definitions) { $map[$definition.Name] = $definition }
    $selected = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $pending = [Collections.Generic.Stack[string]]::new()
    foreach ($target in $Targets) { $pending.Push($target) }
    while ($pending.Count) {
        $name = $pending.Pop()
        if (-not $selected.Add($name)) { continue }
        if ($Overrides.Contains($name)) { continue }
        if (-not $map.ContainsKey($name)) { throw "Missing production fixture target: $name" }
        foreach ($reference in @(Get-MReferences $map[$name].Expression)) {
            if ($map.ContainsKey($reference) -or $Overrides.Contains($reference)) { $pending.Push($reference) }
        }
    }
    foreach ($name in $selected) {
        $token = if ($map.ContainsKey($name)) { $map[$name].Token } else { '#"' + $name.Replace('"','""') + '"' }
        $expression = if ($Overrides.Contains($name)) { $Overrides[$name] } else { $map[$name].Expression }
        $token + ' = ' + $expression
    }
}

$exactSource = Join-Path $repoRoot $SourcePath
$definitions = @(Get-MDefinitions ([IO.File]::ReadAllText($exactSource)))
if ($definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) { throw 'Duplicate production shared query names.' }
$overrides = [ordered]@{
    'Master Source Prepare' = 'Inputs[Master]'
    'INPUT SHiftEnd' = 'Inputs[ShiftEnds]'
    'INPUT MinuteWorkers' = 'Inputs[Roles]'
    'EXTRACT Roles' = 'Inputs[Roles]'
    'EXTRACT RoleAnalysis' = 'Inputs[Analysis]'
    'EXTRACT Units' = 'Inputs[Units]'
    'EXTRACT Unit Analysis' = 'Inputs[UnitAnalysis]'
    'MinuteWorkersFTE_TABLE' = 'Inputs[CareDemand]'
    'MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE' = 'Inputs[CareProfile]'
    'MinuteWorkersFTE_CHECK' = 'Inputs[CareChecks]'
    'MW Publication Check' = 'Inputs[CareChecks]'
}
$bindings = @(Get-Bindings $definitions $overrides @('RoleDemand_TABLE','RoleDemand_CHECK'))
$evaluate = '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + "`nin [Demand=RoleDemand_TABLE, Checks=RoleDemand_CHECK]"
$fixturePath = Join-Path $PSScriptRoot 'fixtures.pq'
$fixture = @(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)) | Where-Object Name -eq 'CSRoleDemandFixtures')
if ($fixture.Count -ne 1) { throw 'Expected one CSRoleDemandFixtures definition.' }
$testSource = "section CSRoleDemandNative;`nshared CSRoleDemandNativeTests = () as table => let`nEvaluate = " + $evaluate + ",`nFixtures = " + $fixture[0].Expression + "`nin Fixtures(Evaluate);`n"
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(|\bValue\.NativeQuery\s*\(') {
    throw 'Synthetic fixtures must not contain external file, workbook, or network calls.'
}
if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) { throw 'Installed Microsoft Power Query SDK Tools are required.' }
$extensionPath = Join-Path $PSScriptRoot 'CSRoleDemandTests.mez'
$queryPath = Join-Path $PSScriptRoot 'fixtures.query.pq'
$resultPath = Join-Path $PSScriptRoot 'native-result.json'
$errorPath = Join-Path $PSScriptRoot 'native-stderr.txt'
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'generated-extension.pq'),$testSource,[Text.UTF8Encoding]::new($false))
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream = [IO.FileStream]::new($extensionPath,[IO.FileMode]::Create)
$archive = [IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
try {
    $entry = $archive.CreateEntry('CSRoleDemandTests.pq')
    $writer = [IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false))
    try { $writer.Write($testSource) } finally { $writer.Dispose() }
} finally { $archive.Dispose(); $stream.Dispose() }
[IO.File]::WriteAllText($queryPath,'CSRoleDemandNativeTests()',[Text.UTF8Encoding]::new($false))
if ($PrepareOnly) {
    [ordered]@{Status='Prepared only';NativeEvaluationPerformed=$false;ExternalConnectorCalls=0;Source=$SourcePath;SourceSHA256=(Get-FileHash -LiteralPath $exactSource).Hash} | ConvertTo-Json
    return
}
$probe = [Text.Encoding]::UTF8.GetBytes('CS-role-demand-native-test')
$encrypted = [Security.Cryptography.ProtectedData]::Protect($probe,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
[void][Security.Cryptography.ProtectedData]::Unprotect($encrypted,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
$info = [Diagnostics.ProcessStartInfo]::new($PQTestPath)
$info.UseShellExecute=$false
$info.CreateNoWindow=$true
$info.WindowStyle='Hidden'
$info.RedirectStandardOutput=$true
$info.RedirectStandardError=$true
foreach ($argument in @('-cfp',(Join-Path $PSScriptRoot 'fixture-credentials.bin'),'run-test','--extension',$extensionPath,'--queryFile',$queryPath,'--prettyPrint')) { $info.ArgumentList.Add($argument) }
$process = [Diagnostics.Process]::Start($info)
try {
    $stdout=$process.StandardOutput.ReadToEndAsync()
    $stderr=$process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(60000)) { $process.Kill($true); throw 'Isolated native fixtures exceeded 60 seconds.' }
    $output=$stdout.GetAwaiter().GetResult()
    [IO.File]::WriteAllText($resultPath,$output)
    [IO.File]::WriteAllText($errorPath,$stderr.GetAwaiter().GetResult())
    $jsonStart=$output.IndexOf('[')
    if ($jsonStart -lt 0) { throw 'No native result; inspect native-result.json and native-stderr.txt.' }
    $results=@($output.Substring($jsonStart) | ConvertFrom-Json)
    $failedQueries=@($results | Where-Object Status -ne 'Passed')
    $failedRows=@($results | ForEach-Object { @($_.Output | Where-Object Passed -ne $true) })
    $dataSources=@($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
    $summary=[ordered]@{Status=if($process.ExitCode -eq 0 -and $results.Count -and -not $failedQueries.Count -and -not $failedRows.Count -and -not $dataSources){'Passed'}else{'Failed'};FixtureRows=[int]($results|Measure-Object RowCount -Sum).Sum;FailedRows=$failedRows.Count;DataSources=$dataSources;Source=$SourcePath;SourceSHA256=(Get-FileHash -LiteralPath $exactSource).Hash;Scope='Actual RoleDemand production expressions with synthetic boundary tables; zero workbook/file/network sources'}
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'summary.json'),($summary | ConvertTo-Json -Depth 6))
    $summary | ConvertTo-Json -Depth 6
    if ($summary.Status -ne 'Passed') { $results | ConvertTo-Json -Depth 12; throw 'CS role demand native fixtures failed.' }
} finally { $process.Dispose() }
