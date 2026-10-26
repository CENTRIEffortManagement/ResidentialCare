[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$helperPath = Join-Path $repoRoot 'outputs/cs-inclusion-bd-step1/run-native-fixtures.ps1'
$helperTokens=$null
$helperErrors=$null
$helperAst=[Management.Automation.Language.Parser]::ParseFile($helperPath,[ref]$helperTokens,[ref]$helperErrors)
if ($helperErrors.Count) { throw 'Scratch M-definition helper has PowerShell parse errors.' }
foreach ($helperName in @('Get-MDefinitions','Get-MReferences','Get-Bindings')) {
    $function=@($helperAst.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object Name -eq $helperName)
    if ($function.Count -ne 1) { throw "Missing scratch helper: $helperName" }
    . ([scriptblock]::Create($function[0].Extent.Text))
}
$sourceRelative='CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/2-DemandExtract.xlsx_PowerQuery.m'
$sourcePath=Join-Path $repoRoot $sourceRelative
$definitions=@(Get-MDefinitions ([IO.File]::ReadAllText($sourcePath)))
if ($definitions.Count -ne @($definitions.Name | Select-Object -Unique).Count) { throw 'Duplicate source query names.' }
$pathDefinition=@($definitions | Where-Object Name -eq 'UnitL1PathTABLE')
if ($pathDefinition.Count -ne 1) { throw 'Expected the authoritative path resolver.' }
$pathLookup='Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content]'
if (-not $pathDefinition[0].Expression.Contains($pathLookup)) { throw 'The exact path input injection boundary changed.' }
$pathDefinition[0].Expression=$pathDefinition[0].Expression.Replace($pathLookup,'Inputs[PathInput]')
$overrides=[ordered]@{
    'IMPORT CentriSyncPaths'='Inputs[PathNavigation]'
    'IMPORT Distributed Demand'='Inputs[DemandNavigation]'
    'IMPORT Settings Data'='Inputs[Settings]'
    'LINK Roles'='Inputs[Roles]'
    'LINK RoleAnalysis'='Inputs[Analysis]'
    'LINK Facility Analysis'='Inputs[Facilities]'
    'IMPORT Role Lists'='error "SharePoint access is disabled in synthetic fixtures."'
    'IMPORT Whiddon Facility Lists'='error "SharePoint access is disabled in synthetic fixtures."'
}
$targets=@('ShiftUnitDemandHRS','DemandExtraction_CHECK','DemandExtraction_RECONCILIATION',
    'Distributed FTE Prepare','Distributed FTE Source Cells','Distributed FTE Profile Prepare','Distributed FTE Exclusions',
    'Demand Calendar CHECK','Demand Roles','Demand Facility','Unit1Path')
$bindings=@(Get-Bindings $definitions $overrides $targets)
$evaluate='(Inputs as record) as record => let' + "`n" + ($bindings -join ",`n") + @'

in [Demand=ShiftUnitDemandHRS,Checks=DemandExtraction_CHECK,Reconciliation=DemandExtraction_RECONCILIATION,
    Pattern=#"Distributed FTE Prepare",Cells=#"Distributed FTE Source Cells",Profile=#"Distributed FTE Profile Prepare",
    Exclusions=#"Distributed FTE Exclusions",CalendarCheck=#"Demand Calendar CHECK",
    Roles=#"Demand Roles",Facility=#"Demand Facility",UnitPath=Unit1Path]
'@
$fixturePath=Join-Path $PSScriptRoot 'fixtures.pq'
$fixture=@(Get-MDefinitions ([IO.File]::ReadAllText($fixturePath)) | Where-Object Name -eq 'CSUnitDemandFixtures')
if ($fixture.Count -ne 1) { throw 'Expected one CSUnitDemandFixtures function.' }
$testSource="section CSUnitDemandNative;`nshared CSUnitDemandNativeTests = () as table => let`nEvaluate = " + $evaluate + ",`nFixtures = " + $fixture[0].Expression + "`nin Fixtures(Evaluate);`n"
if ($testSource -match '(?i)\b(File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(|\bValue\.NativeQuery\s*\(') {
    throw 'Prepared native package still contains external source calls.'
}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'generated-extension.pq'),$testSource,[Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'fixtures.query.pq'),'CSUnitDemandNativeTests()',[Text.UTF8Encoding]::new($false))
Add-Type -AssemblyName System.IO.Compression.FileSystem
$stream=[IO.FileStream]::new((Join-Path $PSScriptRoot 'CSUnitDemandTests.mez'),[IO.FileMode]::Create)
$archive=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Create)
try {
    $entry=$archive.CreateEntry('CSUnitDemandTests.pq')
    $writer=[IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false))
    try { $writer.Write($testSource) } finally { $writer.Dispose() }
} finally { $archive.Dispose();$stream.Dispose() }
$summary=[ordered]@{Status='Prepared only';Source=$sourceRelative;SourceSHA256=(Get-FileHash -LiteralPath $sourcePath).Hash;
    NativeEvaluationPerformed=$false;ExternalConnectorCalls=0;
    FixtureCases=([regex]::Matches([IO.File]::ReadAllText($fixturePath),'(?m)^\s*Check\("')).Count;
    Scope='Actual BD Demand Extract expressions with synthetic handover and Settings navigation; SDK not launched'}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'preparation-summary.json'),($summary | ConvertTo-Json -Depth 4))
$summary | ConvertTo-Json -Depth 4
