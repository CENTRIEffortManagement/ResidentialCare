[CmdletBinding()]
param([switch] $PrepareOnly)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$sourcePath = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m'
$tokens=$null; $errors=$null
$parser = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'),[ref]$tokens,[ref]$errors)
$definition = @($parser.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object Name -eq 'Read-MDefinitions')
if ($errors.Count -or $definition.Count -ne 1) { throw 'Approved M parser is unavailable.' }
. ([scriptblock]::Create($definition[0].Extent.Text))
$queries = Read-MDefinitions $sourcePath
$selected = @('AvailabilityText','LINK Facilities','FacilityMappings_Prepare','ReconciledWorkers_Facility_Prepare','AvailabilityLeaveSource_Prepare','AllocationRoleHours_Prepare')
$bindings = @($selected | ForEach-Object {
    if (-not $queries.Contains($_)) { throw "Missing source query: $_" }
    '#"' + $_.Replace('"','""') + '" = ' + $queries[$_]
})
$prefix = @'
section FacilityMappingTests;
shared FacilityMappingTests = () as table => let
    Facilities = #table({"Title","Facility-Abbrev","FacilityShort"},{
        {"BD: Beaudesert","BD","BD"}, {"JE: Jessie Hunt","JH-RY","JE"},
        {"NR: Narrabri- Robert Young","JH-RY","NR"}, {"TE: Temora","TE","TE"},
        {"TE: Temora Greenstone","TE","TE"}, {"TE: Temora","TE","TE"}}),
    Workers = #table({"EmployeeID","Facility-Abbrev","Role"}, {
        {"JE1","JE","AIN"},{"NR1","NR","AIN"},{"CAN1","JH-RY","RN"},
        {"BD1","BD","EN"},{"TE1","TE","RN"},{"XX1","XX","AIN"},{"NULL1",null,"AIN"}}),
    EventSource = #table({"Column1","Column2","Column3","Column4","Column5","Column6","Column7"},{
        {"Site Name","Employee Name","PayrollCode","Type","DateTimeStart","DateTimeEnd","Comments"},
        {"NR: Narrabri- Robert Young","Worker NR","NR1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null},
        {"JE: Jessie Hunt","Worker JE","JE1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null},
        {"TE: Temora","Worker TE","TE1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null},
        {"BD: Beaudesert","Worker BD","BD1","AVAIL",#datetime(2026,10,1,0,0,0),#datetime(2026,10,2,0,0,0),null}}),
    AllocationSource = #table({"Location","Code","Role","Hours","Name"},{
        {"NR: Narrabri- Robert Young","CROSS1","AIN",2,"Worker Cross"},
        {"JE: Jessie Hunt","CROSS1","AIN",3,"Worker Cross"},
        {"TE: Temora","TE1","RN",4,"Worker TE"},{"BD: Beaudesert","BD1","EN",5,"Worker BD"}}),
    Evaluate = (reference as table, selectedUnit as text, sourceWorkers as table, enabled as logical) as record => let
        Unit = selectedUnit,
        #"IMPORT Facilities List" = #table({"Id","Items"},{{"1987eadb-ad2e-491a-a927-e5585667d4c5",reference}}),
        ReconciledWorkerMembership_Prepare = sourceWorkers,
        #"IMPORT Availability Leave Source" = EventSource,
        #"IMPORT AllocationExtracted" = AllocationSource,
        FacilityAnalysisTABLE = if enabled then #table({"Unit"},{{selectedUnit}}) else #table({"Unit"},{}),
        AvailabilityRoleGroups = {"AIN","EN","RN"},
'@
$suffix = @'
    in [Raw=#"LINK Facilities",Mapping=FacilityMappings_Prepare,Workers=ReconciledWorkers_Facility_Prepare,
        Events=AvailabilityLeaveSource_Prepare,Allocations=AllocationRoleHours_Prepare],
    Baseline = Evaluate(Facilities,"JH-RY",Workers,true),
    Worker = (id as text) as record => Table.SelectRows(Baseline[Workers],each [EmployeeID]=id){0},
    Check = (name as text, run as function) as record => let Attempt=try run(),
        Passed=if Attempt[HasError] then false else Attempt[Value]=true
    in [Case=name,Passed=Passed,Details=if Attempt[HasError] then Attempt[Error][Message] else if Passed then null else "Unexpected result"],
    LowerCaseWorkers = #table({"EmployeeID","Facility-Abbrev","Role"},{{"lower"," nr ","AIN"}}),
    ChangedReference = Table.ReplaceValue(Facilities,"NR","ZZ",Replacer.ReplaceValue,{"FacilityShort"}),
    ChangedWorkers = #table({"EmployeeID","Facility-Abbrev","Role"},{{"dynamic","ZZ","AIN"}}),
    ConflictingReference = Table.Combine({Facilities,#table({"Title","Facility-Abbrev","FacilityShort"},{{"Conflicting site","BD","NR"}})}),
    MissingCanonical = Table.ReplaceValue(Facilities,"JH-RY",null,Replacer.ReplaceValue,{"Facility-Abbrev"}),
    LegacyReference = Table.RenameColumns(Facilities,{{"Facility-Abbrev","field_1"}}),
    MatchedEvents = Table.NestedJoin(Baseline[Events],{"Payroll Code","Facility-Abbrev"},Baseline[Workers],{"EmployeeID","Facility-Abbrev"},"Worker",JoinKind.LeftOuter),
    Cases = {
        Check("NR maps to JH-RY",()=> Worker("NR1")[#"Facility-Abbrev"]="JH-RY"),
        Check("JE maps to JH-RY",()=> Worker("JE1")[#"Facility-Abbrev"]="JH-RY"),
        Check("Already canonical code retained",()=> Worker("CAN1")[#"Facility-Abbrev"]="JH-RY"),
        Check("BD preserved",()=> Worker("BD1")[#"Facility-Abbrev"]="BD"),
        Check("TE preserved",()=> Worker("TE1")[#"Facility-Abbrev"]="TE"),
        Check("Original NR code remains visible",()=> Worker("NR1")[SourceFacility]="NR"),
        Check("Unmapped value remains diagnostic",()=> Worker("XX1")[#"Facility-Abbrev"]="XX" and Worker("XX1")[FacilityMappingStatus]="Unmapped facility: inspect LINK Facilities"),
        Check("Null facility remains diagnostic",()=> Worker("NULL1")[#"Facility-Abbrev"]=null and Worker("NULL1")[FacilityMappingStatus]="Missing facility"),
        Check("Worker row grain preserved",()=> Table.RowCount(Baseline[Workers])=Table.RowCount(Workers)),
        Check("Duplicate TE mappings do not multiply lookup",()=> Table.RowCount(Baseline[Mapping])=10),
        Check("Lookup keys unique",()=> Table.RowCount(Baseline[Mapping])=Table.RowCount(Table.Distinct(Baseline[Mapping],{"FacilityKey"}))),
        Check("Case and whitespace normalized",()=> Evaluate(Facilities,"JH-RY",LowerCaseWorkers,true)[Workers]{0}[#"Facility-Abbrev"]="JH-RY"),
        Check("Reference-driven alias without code edits",()=> Evaluate(ChangedReference,"JH-RY",ChangedWorkers,true)[Workers]{0}[#"Facility-Abbrev"]="JH-RY"),
        Check("NR and JE availability included",()=> Table.RowCount(Baseline[Events])=2 and List.AllTrue(List.Transform(Baseline[Events][#"Facility-Abbrev"],each _="JH-RY"))),
        Check("NR and JE source identities join to workers",()=> List.AllTrue(List.Transform(MatchedEvents[Worker],each Table.RowCount(_)=1))),
        Check("Both sites allocation grouped once",()=> Table.RowCount(Baseline[Allocations])=1 and Baseline[Allocations]{0}[AllocatedHours]=5 and Baseline[Allocations]{0}[#"Facility-Abbrev"]="JH-RY"),
        Check("BD unit scope preserved",()=> Table.RowCount(Evaluate(Facilities,"BD",Workers,true)[Events])=1 and Evaluate(Facilities,"BD",Workers,true)[Allocations]{0}[AllocatedHours]=5),
        Check("TE duplicates do not multiply events or hours",()=> Table.RowCount(Evaluate(Facilities,"TE",Workers,true)[Events])=1 and Evaluate(Facilities,"TE",Workers,true)[Allocations]{0}[AllocatedHours]=4),
        Check("Disabled analysis remains excluded",()=> Table.IsEmpty(Evaluate(Facilities,"JH-RY",Workers,false)[Events]) and Table.IsEmpty(Evaluate(Facilities,"JH-RY",Workers,false)[Allocations])),
        Check("Legacy field_1 projection supported",()=> Evaluate(LegacyReference,"JH-RY",Workers,true)[Workers]{1}[#"Facility-Abbrev"]="JH-RY"),
        Check("LINK lands full list unchanged",()=> Table.ColumnNames(Evaluate(Table.AddColumn(Facilities,"Extra",each 1),"JH-RY",Workers,true)[Raw])=Table.ColumnNames(Table.AddColumn(Facilities,"Extra",each 1)) and Table.RowCount(Baseline[Raw])=6),
        Check("Conflicting alias blocked",()=> (try Table.RowCount(Evaluate(ConflictingReference,"JH-RY",Workers,true)[Mapping]))[HasError]),
        Check("Missing FacilityShort is explicit error",()=> (try Table.RowCount(Evaluate(Table.RemoveColumns(Facilities,{"FacilityShort"}),"JH-RY",Workers,true)[Mapping]))[HasError]),
        Check("Missing canonical facility blocked",()=> (try Table.RowCount(Evaluate(MissingCanonical,"JH-RY",Workers,true)[Mapping]))[HasError])
    }
in Table.FromRecords(Cases);
'@
$testSource = $prefix + ($bindings -join ",`n") + "`n" + $suffix
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(') { throw 'Synthetic tests contain external access.' }
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'fixtures.pq'),$testSource,[Text.UTF8Encoding]::new($false))
if ($PrepareOnly) { 'Prepared 24 native synthetic cases from the actual edited M queries. No Excel or source data accessed.'; return }
$pqTestPath = Join-Path ([IO.Path]::GetTempPath()) 'ResidentialCare-StaffList-PQTest-2.155.2/package/tools/PQTest.exe'
if (-not (Test-Path -LiteralPath $pqTestPath)) { throw 'Installed Microsoft Power Query test runtime is required.' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$extensionPath=Join-Path $PSScriptRoot 'FacilityMappingTests.mez'
if (Test-Path -LiteralPath $extensionPath) { [IO.File]::Delete($extensionPath) }
$archive=[IO.Compression.ZipFile]::Open($extensionPath,[IO.Compression.ZipArchiveMode]::Create)
try { $entry=$archive.CreateEntry('FacilityMappingTests.pq'); $writer=[IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false)); try {$writer.Write($testSource)} finally {$writer.Dispose()} } finally {$archive.Dispose()}
$queryPath=Join-Path $PSScriptRoot 'fixtures.query.pq'
[IO.File]::WriteAllText($queryPath,'FacilityMappingTests()',[Text.UTF8Encoding]::new($false))
$info=[Diagnostics.ProcessStartInfo]::new($pqTestPath)
$info.UseShellExecute=$false; $info.CreateNoWindow=$true; $info.WindowStyle='Hidden'
$info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
foreach($argument in @('-cfp',(Join-Path $PSScriptRoot 'fixture-credentials.bin'),'run-test','--extension',$extensionPath,'--queryFile',$queryPath,'--prettyPrint')) {$info.ArgumentList.Add($argument)}
$process=[Diagnostics.Process]::Start($info)
try {
    $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
    if(-not $process.WaitForExit(60000)) {$process.Kill($true);throw 'Synthetic tests exceeded 60 seconds.'}
    $output=$stdout.GetAwaiter().GetResult()
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'native-result.json'),$output)
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot 'native-stderr.txt'),$stderr.GetAwaiter().GetResult())
    $start=$output.IndexOf('[')
    if($start -lt 0) {throw 'Native test engine returned no JSON; inspect native-result.json and native-stderr.txt.'}
    $results=@($output.Substring($start)|ConvertFrom-Json)
    $failed=@($results|Where-Object Status -ne 'Passed')
    $failedRows=@($results|ForEach-Object {@($_.Output|Where-Object Passed -ne $true)})
    $sources=@($results|ForEach-Object {@($_.DataSourceAnalysis)}).Count
    $rowCount=[int]($results|Measure-Object RowCount -Sum).Sum
    if($process.ExitCode -ne 0 -or -not $results.Count -or $failed.Count -or $failedRows.Count -or $sources -or $rowCount -ne 24) {
        $results|ConvertTo-Json -Depth 12;throw 'Native facility-mapping cases failed.'
    }
    $summary=[ordered]@{Status='Passed';FixtureRows=$rowCount;FailedRows=0;DataSources=$sources;SourceSHA256=(Get-FileHash -LiteralPath $sourcePath).Hash;Scope='Actual JH-RY facility lookup and worker/event/allocation queries, with synthetic Facilities, workers, events and allocation only'}
    $summary|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'native-summary.json') -Encoding utf8
    $summary|ConvertTo-Json -Depth 4
} finally {$process.Dispose()}
