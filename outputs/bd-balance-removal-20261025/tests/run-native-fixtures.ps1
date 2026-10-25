[CmdletBinding()]
param(
    [string] $CapacitySource = 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m',
    [string] $MembershipBlock,
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-StaffList-PQTest-2.155.2/package/tools/PQTest.exe'),
    [switch] $PrepareOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path

function Get-MDefinitions([string] $Text) {
    foreach ($match in [regex]::Matches($Text, '(?m)^shared\s+(?<name>#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=')) {
        $start = $match.Index + $match.Length
        $inString = $false; $lineComment = $false; $end = -1
        for ($index = $start; $index -lt $Text.Length; $index++) {
            $character = $Text[$index]
            $next = if ($index + 1 -lt $Text.Length) { $Text[$index + 1] } else { [char]0 }
            if ($lineComment) { if ($character -eq "`n") { $lineComment = $false }; continue }
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
        [pscustomobject]@{ Name=$name; Token=$token; Expression=$Text.Substring($start,$end-$start).Trim() }
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

$sourcePath = Join-Path $repoRoot $CapacitySource
$definitions = @(Get-MDefinitions ([IO.File]::ReadAllText($sourcePath)))
if ($definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) { throw 'Duplicate production shared query names.' }
if ($MembershipBlock) {
    $blockDefinitions = @(Get-MDefinitions ([IO.File]::ReadAllText((Join-Path $repoRoot $MembershipBlock))))
    $definitions = @($definitions | Where-Object Name -notin $blockDefinitions.Name) + $blockDefinitions
}
$overrides = [ordered]@{
    'IMPORT Published Roster' = 'Inputs[Roster]'
    'IMPORT Availability Leave Source' = 'Inputs[Availability]'
    "IMPORT Worker's Report" = 'Inputs[HR]'
    'LINK Facilities' = 'Inputs[Facilities]'
    'LINK Roles' = '#table({"Roster Roles","Role Group"},{{"Assistant in Nursing AM","AIN"},{"Enrolled Nurse AM","EN"},{"Registered Nurse AM","RN"},{"REGN AM","RN"}})'
    'FacilityAnalysisTABLE' = '#table({"Facility-Abbrev"},{{Inputs[Unit]}})'
    'Unit' = 'Inputs[Unit]'
    'AllocationRoleShares' = '#table(type table [EmployeeID=text, #"Facility-Abbrev"=text, Role=text, RoleCount=number, RoleShare=number],{})'
}
$targets = @('ReconciledWorkerMembership_Prepare','ReconciledWorkers_Prepare','ReconciledWorkers_Eligible','ReconciledWorkers_CHECK')
$bindings = @(Get-Bindings $definitions $overrides $targets)
$evaluate = '(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + @'

in [Membership=ReconciledWorkerMembership_Prepare, Prepared=ReconciledWorkers_Prepare,
    Eligible=ReconciledWorkers_Eligible, Checks=ReconciledWorkers_CHECK]
'@

$staffPath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/StaffListMaster.xlsx_PowerQuery.m'
$staffDefinitions = @(Get-MDefinitions ([IO.File]::ReadAllText($staffPath)))
$staffOverrides = [ordered]@{
    'Masterlist_Base' = 'Inputs[Staff]'
    'EXTRACT Availability Staff List' = 'Inputs[Staff]'
    'EXTRACT StaffList RosterDates' = 'Table.FromColumns({List.Dates(#date(2026,10,1),29,#duration(1,0,0,0))},{"Date"})'
    'EXTRACT StaffList ShiftDuration' = '7.6'
    'EXTRACT StaffList MaxAvailability' = '40'
}
$staffBindings = @(Get-Bindings $staffDefinitions $staffOverrides @('Masterlist','ResourceContract_CHECK'))
$staffEvaluate = '(Inputs as record) as record => let' + "`n" + ($staffBindings -join ",`n") + @'

in [Published=Masterlist, Checks=ResourceContract_CHECK]
'@

$fixturePath = Join-Path $PSScriptRoot 'fixtures.pq'
$fixtureDefinition = @(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)) | Where-Object Name -eq 'BDBalanceRemovalFixtures')
if ($fixtureDefinition.Count -ne 1) { throw 'Expected one fixture definition.' }
$testSource = "section BDBalanceRemovalNative;`nshared BDBalanceRemovalTests = () as table => let`nEvaluate = " + $evaluate + ",`nEvaluateStaff = " + $staffEvaluate + ",`nFixtures = " + $fixtureDefinition[0].Expression + "`nin Fixtures(Evaluate,EvaluateStaff);`n"
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(|\bValue\.NativeQuery\s*\(') {
    throw 'Generated synthetic fixtures must not access external sources.'
}
if (-not (Test-Path -LiteralPath $PQTestPath -PathType Leaf)) { throw 'Installed Microsoft Power Query SDK Tools are required.' }
$extensionPath = Join-Path $PSScriptRoot 'BDBalanceRemovalTests.mez'
$queryPath = Join-Path $PSScriptRoot 'fixtures.query.pq'
$resultPath = Join-Path $PSScriptRoot 'native-result.json'
$errorPath = Join-Path $PSScriptRoot 'native-stderr.txt'
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'generated-extension.pq'),$testSource,[Text.UTF8Encoding]::new($false))
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path -LiteralPath $extensionPath) { [IO.File]::Delete($extensionPath) }
$archive = [IO.Compression.ZipFile]::Open($extensionPath,[IO.Compression.ZipArchiveMode]::Create)
try {
    $entry = $archive.CreateEntry('BDBalanceRemovalTests.pq')
    $writer = [IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false))
    try { $writer.Write($testSource) } finally { $writer.Dispose() }
} finally { $archive.Dispose() }
[IO.File]::WriteAllText($queryPath,'BDBalanceRemovalTests()',[Text.UTF8Encoding]::new($false))
if ($PrepareOnly) {
    [ordered]@{Status='Prepared only'; ExternalConnectorCalls=0; NativeEvaluationPerformed=$false;
        CapacitySHA256=(Get-FileHash -LiteralPath $sourcePath).Hash;
        StaffSHA256=(Get-FileHash -LiteralPath $staffPath).Hash} | ConvertTo-Json
    return
}
# The SDK requires Windows profile encryption; preparation above never launches it.
$profileProbe = [Text.Encoding]::UTF8.GetBytes('BD-Balance-removal-native-test')
$encryptedProbe = [Security.Cryptography.ProtectedData]::Protect($profileProbe,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
[void][Security.Cryptography.ProtectedData]::Unprotect($encryptedProbe,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
$info = [Diagnostics.ProcessStartInfo]::new($PQTestPath)
$info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.WindowStyle='Hidden'
$info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
foreach ($argument in @('-cfp',(Join-Path $PSScriptRoot 'fixture-credentials.bin'),'run-test','--extension',$extensionPath,'--queryFile',$queryPath,'--prettyPrint')) { $info.ArgumentList.Add($argument) }
$process = [Diagnostics.Process]::Start($info)
try {
    $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(60000)) {
        $process.Kill($true)
        [IO.File]::WriteAllText($resultPath,$stdout.GetAwaiter().GetResult())
        [IO.File]::WriteAllText($errorPath,$stderr.GetAwaiter().GetResult())
        throw 'Isolated native tests exceeded 60 seconds.'
    }
    $output=$stdout.GetAwaiter().GetResult()
    [IO.File]::WriteAllText($resultPath,$output)
    [IO.File]::WriteAllText($errorPath,$stderr.GetAwaiter().GetResult())
    $jsonStart=$output.IndexOf('[')
    if ($jsonStart -lt 0) { throw 'No native result returned; inspect native-result.json and native-stderr.txt.' }
    $results=@($output.Substring($jsonStart) | ConvertFrom-Json)
    $failedQueries=@($results | Where-Object Status -ne 'Passed')
    $failedRows=@($results | ForEach-Object { @($_.Output | Where-Object Passed -ne $true) })
    $dataSources=@($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
    $summary=[ordered]@{Status=if($process.ExitCode -eq 0 -and $results.Count -and -not $failedQueries.Count -and -not $failedRows.Count -and -not $dataSources){'Passed'}else{'Failed'};FixtureRows=[int]($results|Measure-Object RowCount -Sum).Sum;FailedRows=$failedRows.Count;DataSources=$dataSources;CapacitySource=$CapacitySource;CapacitySHA256=(Get-FileHash -LiteralPath $sourcePath).Hash;StaffSHA256=(Get-FileHash -LiteralPath $staffPath).Hash;MembershipBlock=$MembershipBlock;Scope='Actual selected production M expressions; synthetic tables only; no workbook/file/network data access'}
    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'summary.json') -Encoding utf8
    $summary | ConvertTo-Json -Depth 6
    if ($summary.Status -ne 'Passed') { $results | ConvertTo-Json -Depth 10; throw 'Native Balance-removal fixtures failed.' }
} finally { $process.Dispose() }
