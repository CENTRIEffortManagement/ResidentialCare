[CmdletBinding()]
param(
    [string] $PQTestPath = (Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-StaffList-PQTest-2.155.2/package/tools/PQTest.exe')
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent

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
        [pscustomobject]@{ Token=$match.Groups['name'].Value; Expression=$Text.Substring($start,$end-$start).Trim() }
    }
}

$selected = @('#"EXTRACT StaffList RosterDates"','ResourceContract','ResourceContract_CHECK','StaffListCheckResult','Masterlist')
$evaluators = [Collections.Generic.List[string]]::new()
$sourceHashes = [ordered]@{}
$referenceBody = $null
foreach ($unitId in @('BD','TE','JH-RY')) {
    $sourcePath = Join-Path $repoRoot "CLIENT/DATExx-Whiddon/UNITS/$unitId/2. Calculations/StaffListMaster.xlsx_PowerQuery.m"
    $text = [IO.File]::ReadAllText($sourcePath).Replace("`r`n","`n")
    $body = $text.Substring($text.IndexOf('section Section1;')).TrimEnd()
    if ($null -eq $referenceBody) { $referenceBody = $body }
    elseif ($body -cne $referenceBody) { throw "StaffListMaster query sections differ: $unitId" }
    $sourceHashes[$unitId] = (Get-FileHash -LiteralPath $sourcePath).Hash
    $definitions = @(Get-MDefinitions $text)
    if ($definitions.Count -ne @($definitions.Token | Select-Object -Unique).Count) { throw "Duplicate shared definitions: $unitId" }
    $bindings = @($definitions | Where-Object Token -in $selected | ForEach-Object { $_.Token + ' = ' + $_.Expression })
    if ($bindings.Count -ne $selected.Count) { throw "Missing test definitions: $unitId" }
    $name = 'Evaluate' + ($unitId -replace '-','')
    $evaluators.Add($name + ' = (dates as table, hours as nullable number, shiftHours as number, maximum as number) as record => let ' + @'
        Base = [Name="Synthetic worker", Role="AIN", Resource=1, Misalignment=0, Source="Availability",
            EmployeeID="fixture", #"Facility-Abbrev"="fixture-unit", EmploymentType=if hours=null then "Casual" else "Permanent",
            PreferredRole="AIN", #"In Allocation"=true, #"In Records"=true, #"Worker Membership"="Both",
            #"Worker Record Status"="Found", #"Worker Contract Issue"=null, #"Contracted FN Hours"=hours,
            RosterShiftCap=if hours=null then 13 else null, AppliedCasualRosterShiftCap=if hours=null then 13 else null, CasualCapIssue=null],
        Masterlist_Base = Table.FromRecords({Base}),
        #"EXTRACT Availability Staff List" = Masterlist_Base,
        #"EXTRACT StaffList ShiftDuration" = shiftHours,
        #"EXTRACT StaffList MaxAvailability" = maximum,
        #"IMPORT StaffList Settings" = #table({"Item","Kind","Data"},{{"PermutationDimensions","Table",dates}}),
'@ + ($bindings -join ",`n") + @'
        , Published = Masterlist,
        Result = Published{0}
    in Record.Combine({Result,[PublishedCount=Table.RowCount(Published),DateCount=Table.RowCount(#"EXTRACT StaffList RosterDates") ]})
'@)
}
$fixture = @'
    DateTable = (days as number) as table => Table.FromColumns({List.Dates(#date(2026,9,28),days,#duration(1,0,0,0))},{"Date"}),
    Check = (name as text, run as function, predicate as function) as record => let
        Attempt=try run(), Passed=if Attempt[HasError] then false else (try predicate(Attempt[Value]) otherwise false)
    in [Case=name,Passed=Passed,Details=if Attempt[HasError] then (try Attempt[Error][Message] otherwise "Evaluation failed") else if Passed then null else "Unexpected result"],
    Evaluators = {[Unit="BD",Evaluate=EvaluateBD],[Unit="TE",Evaluate=EvaluateTE],[Unit="JH-RY",Evaluate=EvaluateJHRY]},
    Cases = List.Combine(List.Transform(Evaluators,(unit) =>
        List.Transform({7,14,21,28,29,35,42},(days) => Check(unit[Unit]&" contract "&Number.ToText(days)&" days",
            () => unit[Evaluate](DateTable(days),76,7.6,40), (result) => result[PublishedCount]=1 and result[DateCount]=days
                and Number.Abs(result[#"Roster Fortnights"]-days/14)<0.00000001
                and Number.Abs(result[#"Contracted Roster Hours"]-76*days/14)<0.00000001
                and result[#"Contracted Shifts"]=Number.RoundDown(76*(days/14)/7.6)
                and result[#"Effective Shift Cap"]=List.Min({40,Number.RoundDown(76*(days/14)/7.6)}))) &
        List.Transform({28,29},(days) => Check(unit[Unit]&" casual full-roster cap "&Number.ToText(days),
            () => unit[Evaluate](DateTable(days),null,7.6,40), (result) => result[#"Effective Shift Cap"]=13
                and result[#"Contracted FN Hours"]=null and result[#"Contracted Roster Hours"]=null)) & {
            Check(unit[Unit]&" duplicate dates", () => unit[Evaluate](Table.Combine({DateTable(29),DateTable(29)}),76,7.6,40), each [DateCount]=29),
            Check(unit[Unit]&" shift duration and Settings cap", () => unit[Evaluate](DateTable(29),75,7.5,12), each [#"Effective Shift Cap"]=12),
            Check(unit[Unit]&" gap blocked", () => (try unit[Evaluate](Table.RemoveRows(DateTable(29),12,1),76,7.6,40)[PublishedCount])[HasError], each _=true),
            Check(unit[Unit]&" empty calendar blocked", () => (try unit[Evaluate](DateTable(0),76,7.6,40)[PublishedCount])[HasError], each _=true),
            Check(unit[Unit]&" blank date blocked", () => (try unit[Evaluate](#table({"Date"},{{null}}),76,7.6,40)[PublishedCount])[HasError], each _=true),
            Check(unit[Unit]&" missing Date column blocked", () => (try unit[Evaluate](#table({"Day"},{{1}}),76,7.6,40)[PublishedCount])[HasError], each _=true)
        })),
    Output = Table.FromRecords(Cases)
in Output;
'@
$testSource = "section StaffListDynamicRoster;`nshared StaffListDynamicRosterTests = () as table => let`n" + ($evaluators -join ",`n") + ",`n" + $fixture
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(') {
    throw 'Native fixtures must not access external sources.'
}
if (-not (Test-Path -LiteralPath $PQTestPath)) { throw 'Installed Microsoft Power Query SDK Tools are required.' }
$outputFolder = Join-Path $repoRoot 'outputs/stafflist-dynamic-roster-20261022/native-tests'
[void][IO.Directory]::CreateDirectory($outputFolder)
$extensionPath = Join-Path $outputFolder 'StaffListDynamicRosterTests.mez'
$queryPath = Join-Path $outputFolder 'fixtures.query.pq'
$resultPath = Join-Path $outputFolder 'native-result.json'
$errorPath = Join-Path $outputFolder 'native-stderr.txt'
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path -LiteralPath $extensionPath) { [IO.File]::Delete($extensionPath) }
$archive = [IO.Compression.ZipFile]::Open($extensionPath,[IO.Compression.ZipArchiveMode]::Create)
try {
    $entry = $archive.CreateEntry('StaffListDynamicRosterTests.pq')
    $writer = [IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false))
    try { $writer.Write($testSource) } finally { $writer.Dispose() }
} finally { $archive.Dispose() }
[IO.File]::WriteAllText($queryPath,'StaffListDynamicRosterTests()',[Text.UTF8Encoding]::new($false))
$info = [Diagnostics.ProcessStartInfo]::new($PQTestPath)
$info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.WindowStyle='Hidden'
$info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
foreach ($argument in @('-cfp',(Join-Path $outputFolder 'fixture-credentials.bin'),'run-test','--extension',$extensionPath,'--queryFile',$queryPath,'--prettyPrint')) { $info.ArgumentList.Add($argument) }
$process = [Diagnostics.Process]::Start($info)
try {
    $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(60000)) { $process.Kill($true); throw 'Isolated native roster tests exceeded 60 seconds.' }
    $output=$stdout.GetAwaiter().GetResult()
    [IO.File]::WriteAllText($resultPath,$output)
    [IO.File]::WriteAllText($errorPath,$stderr.GetAwaiter().GetResult())
    $jsonStart=$output.IndexOf('[')
    if ($jsonStart -lt 0) { throw "No native result returned; inspect $resultPath and $errorPath." }
    $results=@($output.Substring($jsonStart) | ConvertFrom-Json)
    $failedQueries=@($results | Where-Object Status -ne 'Passed')
    $failedRows=@($results | ForEach-Object { @($_.Output | Where-Object Passed -ne $true) })
    $dataSources=@($results | ForEach-Object { @($_.DataSourceAnalysis) }).Count
    if ($process.ExitCode -ne 0 -or -not $results.Count -or $failedQueries.Count -or $failedRows.Count -or $dataSources) {
        $results | ConvertTo-Json -Depth 12
        throw 'Native StaffListMaster roster fixtures failed.'
    }
    $summary=[ordered]@{Status='Passed';FixtureRows=[int]($results | Measure-Object RowCount -Sum).Sum;FailedRows=0;DataSources=0;SourceSHA256=$sourceHashes;Scope='Identical BD/TE/JH-RY source; actual date validation, contract proration, required checks and publication; synthetic data only'}
    $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $outputFolder 'summary.json') -Encoding utf8
    $summary | ConvertTo-Json -Depth 5
} finally { $process.Dispose() }
