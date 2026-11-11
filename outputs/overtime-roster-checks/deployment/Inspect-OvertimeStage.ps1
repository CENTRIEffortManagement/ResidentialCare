[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateSet('Roster','Enrichment','Physical','Chronology','Parts','Fortnight','Excess','ShiftResults')] [string] $Stage,
    [ValidateRange(15,180)] [int] $TimeoutSeconds=25,
    [string] $RunFolder='', [string] $LogPath='', [string] $StopRequestPath='', [string] $WorkerStatePath=''
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if(-not $RunFolder){$RunFolder=(Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/deployment-current.txt') -Raw).Trim()}
if([IO.Path]::IsPathRooted($RunFolder)){throw 'RunFolder must be repository-relative.'}
$runRoot=[IO.Path]::GetFullPath((Join-Path $repoRoot $RunFolder))
if(-not $runRoot.StartsWith(($repoRoot.TrimEnd('\','/')+[IO.Path]::DirectorySeparatorChar),[StringComparison]::OrdinalIgnoreCase)){throw 'Diagnostic evidence must stay inside the repository.'}
if(-not $LogPath){$LogPath=Join-Path $runRoot "stage-$Stage.log"}
if(-not $StopRequestPath){$StopRequestPath=Join-Path $runRoot 'stop.txt'}
$workbookRelative='CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx'
$sourceRelative='Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq'
$workbookPath=Join-Path $repoRoot $workbookRelative
$sourcePath=Join-Path $repoRoot $sourceRelative
function Write-Log([string]$Message,[string]$Level='INFO'){$line="[$([DateTime]::UtcNow.ToString('o'))] [$Level] $Stage $Message";Add-Content -LiteralPath $LogPath -Value $line;Write-Host $line}
function Write-CurrentStatus([string]$State,[string]$Message=''){}
function Test-StopNowRequested{(Test-Path -LiteralPath $StopRequestPath) -and ((Get-Content -LiteralPath $StopRequestPath -Raw) -match '(?im)^mode\s*=\s*Now\s*$')}
function Assert-NotStopNowRequested{if(Test-StopNowRequested){throw 'Operator stop observed.'}}
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
if(-not $WorkerStatePath){exit (Invoke-SupervisedExcelRefresh -WorkerScript $PSCommandPath -Parameters @{Stage=$Stage;RunFolder=$RunFolder;LogPath=$LogPath;StopRequestPath=$StopRequestPath;TimeoutSeconds=$TimeoutSeconds} -TimeoutSeconds $TimeoutSeconds)}
$script:WorkerStatePath=$WorkerStatePath
$script:RefreshState=@{phase='Starting';detail=$null;excel=$null;saved=$false;updatedAt=''}
$tokens=$null;$errors=$null
$syncAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Approved synchronization helpers did not parse.'}
foreach($name in @('Get-Sha256','Normalize-Formula','Read-MDefinitions','Release-ComObject')){
    $definition=@($syncAst.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst]},$false)|Where-Object Name -eq $name)
    if($definition.Count -ne 1){throw "Missing approved helper: $name"};. ([scriptblock]::Create($definition[0].Extent.Text))
}
function Replace-FinalIn([string]$Formula,[string]$Step){
    if($Formula -notmatch ('(?m)^\s*'+[regex]::Escape($Step)+'\s*=')){throw "Requested stage step is absent: $Step"}
    $final=[regex]::Match($Formula,'(?s)\bin\s+(?:#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*$')
    if(-not $final.Success){throw 'Could not identify the actual query final in target.'}
    $Formula.Substring(0,$final.Index)+"in`n    "+$Step
}
$excel=$null;$book=$null;$queries=$null;$query=$null;$sheet=$null;$table=$null;$qt=$null;$destination=$null;$range=$null;$lease=$null
$ownsExcel=$false;$closed=$false;$quit=$false;$timer=[Diagnostics.Stopwatch]::new()
try{
    Set-RefreshPhase 'Preflight'
    $sync=Get-Content -LiteralPath (Join-Path $runRoot 'Core-sync-receipt.json') -Raw|ConvertFrom-Json
    if($sync.Target -cne 'Core' -or $sync.Workbook -cne $workbookRelative -or $sync.Source -cne $sourceRelative -or -not $sync.ReExtractionMatches -or -not $sync.Saved -or -not $sync.Closed){throw 'Core exact-pair synchronization is unverified.'}
    $before=Get-Sha256 $workbookPath;$sourceHash=Get-Sha256 $sourcePath
    if($before -cne $sync.WorkbookSHA256After -or $sourceHash -cne $sync.SourceSHA256){throw 'Core workbook or canonical source differs from its synchronization receipt.'}
    $approved=Read-MDefinitions $sourcePath
    . (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/RefreshRunGate.ps1')
    $lease=Enter-RefreshRunGate -DateRoot (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon') -Exclusive
    if(Test-Path -LiteralPath (Join-Path (Split-Path $workbookPath) '~$OvertimeRosterChecks.xlsx')){throw 'Core workbook has an Excel lock.'}
    $handle=[IO.File]::Open($workbookPath,'Open','ReadWrite','None');$handle.Dispose()
    $existing=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
    Set-RefreshPhase 'CreatingExcel'
    $excel=New-Object -ComObject Excel.Application;Register-RefreshExcel -Excel $excel -PreExistingExcelIds $existing;$ownsExcel=$true
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($workbookPath,0,$false);Wait-RefreshWorkbookIdentity -Workbook $book -ExpectedPath $workbookPath
    $queries=$book.Queries
    foreach($name in $approved.Keys){$native=$null;try{$native=$queries.Item($name);if([string]$native.Name -cne $name -or (Normalize-Formula ([string]$native.Formula)) -cne (Normalize-Formula $approved[$name])){throw "Installed query differs: $name"}}finally{Release-ComObject $native}}
    # Use the actual installed query expression. Only the in-memory diagnostic's
    # final in target changes for the two internal stages; production queries do not.
    $shiftQuery=$null
    try{$shiftQuery=$queries.Item('OvertimeRosterChecks_ShiftPrepare');$shiftFormula=[string]$shiftQuery.Formula}finally{Release-ComObject $shiftQuery}
    $selected=switch($Stage){
        'Roster'{'OvertimeRosterChecks_RosterPrepare'}
        'Enrichment'{Replace-FinalIn $shiftFormula 'RemovedNestedInputs'}
        'Physical'{Replace-FinalIn $shiftFormula 'BufferedMatchShifts'}
        'Chronology'{$shiftFormula}
        'Parts'{'OvertimeRosterChecks_FortnightParts'}
        'Fortnight'{'OvertimeRosterChecks_FortnightResults'}
        'Excess'{'OvertimeRosterChecks_ExcessSegments'}
        'ShiftResults'{'OvertimeRosterChecks_ShiftResults'}
    }
    $formula="let`n    StageValue = ("+$selected+"),`n"+@'
    // Table.Buffer forces every scalar field; nested tables/lists remain shallow.
    Forced = Table.Buffer(StageValue),
    Columns = Table.ColumnNames(Forced),
    GroupColumn = if List.Contains(Columns, "IdentityStatus") then "IdentityStatus" else if List.Contains(Columns, "SourceUnit") then "SourceUnit" else null,
    Counts = if GroupColumn = null then #table({"GroupValue", "RowCount"}, {{"All", Table.RowCount(Forced)}})
        else Table.RenameColumns(Table.Group(Forced, {GroupColumn}, {{"RowCount", each Table.RowCount(_), Int64.Type}}), {{GroupColumn, "GroupValue"}}),
    TextGroups = Table.TransformColumns(Counts, {{"GroupValue", each if _ = null then null else Text.From(_), type nullable text}}),
    GroupName = Table.AddColumn(TextGroups, "GroupBy", each if GroupColumn = null then "All" else GroupColumn, type text),
    ColumnNames = Table.AddColumn(GroupName, "ColumnNames", each Text.Combine(Columns, " | "), type text),
    Output = Table.SelectColumns(ColumnNames, {"GroupBy", "GroupValue", "RowCount", "ColumnNames"})
in
    Output
'@
    $query=$queries.Add('OvertimeStageProbe',$formula)
    $sheet=$book.Worksheets.Add();$sheet.Name='OvertimeStageProbe';$destination=$sheet.Range('A1')
    $table=$sheet.ListObjects.Add(0,'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="OvertimeStageProbe";Extended Properties=""',$null,1,$destination)
    $qt=$table.QueryTable;$qt.CommandType=2;$qt.CommandText=@('SELECT * FROM [OvertimeStageProbe]');$qt.BackgroundQuery=$false;$qt.RefreshOnFileOpen=$false
    Set-RefreshPhase "Evaluating$Stage";$timer.Start()
    if(-not $qt.Refresh($false)){throw 'Stage probe did not confirm foreground completion.'}
    $timer.Stop();$range=$table.DataBodyRange
    if($null -eq $range){throw 'Stage probe returned no aggregate rows.'};$values=$range.Value2
    $rows=@(for($index=1;$index -le $values.GetLength(0);$index++){[ordered]@{GroupBy=[string]$values[$index,1];GroupValue=[string]$values[$index,2];RowCount=[long]$values[$index,3];ColumnNames=[string]$values[$index,4]}})
    $aggregateCount=[long]0
    foreach($aggregate in $rows){$aggregateCount+=[long]$aggregate['RowCount']}
    [ordered]@{Stage=$Stage;EvaluationCompleted=$true;ElapsedSeconds=$timer.Elapsed.TotalSeconds;RowCount=$aggregateCount;Aggregates=$rows;Saved=$false;SourceSHA256=$sourceHash;WorkbookSHA256Before=$before;RecordedAtUTC=[DateTime]::UtcNow.ToString('o')}|
        ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $runRoot "stage-$Stage-evaluated.json") -Encoding utf8
    Set-RefreshPhase 'ClosingWithoutSaving';$book.Close($false);$closed=$true;$excel.Quit();$quit=$true
}catch{
    $script:RefreshState.error=@{Phase=$script:RefreshState.phase;Message=$_.Exception.Message};Write-RefreshWorkerState;Write-Log $_.Exception.Message 'ERROR';throw
}finally{
    if($null -ne $book -and -not $closed){try{$book.Close($false)}catch{Write-Log $_.Exception.Message 'WARN'}}
    if($null -ne $excel -and $ownsExcel -and -not $quit){try{$excel.Quit()}catch{Write-Log $_.Exception.Message 'WARN'}}
    foreach($object in @($range,$qt,$table,$destination,$sheet,$query,$queries,$book,$excel)){try{Release-ComObject $object}catch{Write-Log ("COM release: "+$_.Exception.Message) 'WARN'}}
    if($null -ne $lease){try{$lease.Dispose()}catch{Write-Log ("Lease release: "+$_.Exception.Message) 'WARN'}};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$after=Get-Sha256 $workbookPath
if($after -cne $before -or (Get-Sha256 $sourcePath) -cne $sourceHash){throw 'Unsaved stage probe changed the workbook or canonical source.'}
[ordered]@{Stage=$Stage;ReadOnlySucceeded=$true;Workbook=$workbookRelative;Source=$sourceRelative;SourceSHA256=$sourceHash;WorkbookSHA256Before=$before;WorkbookSHA256After=$after;WorkbookUnchanged=$true;CanonicalSourceUnchanged=$true;ElapsedSeconds=$timer.Elapsed.TotalSeconds;RowCount=$aggregateCount;Aggregates=$rows;Saved=$false;Closed=$true;TimeoutSeconds=$TimeoutSeconds;WithinFifteenSeconds=($timer.Elapsed.TotalSeconds -le 15);SupervisorRequiresSaved=$true;ExpectedSupervisorExitCode=1;RecordedAtUTC=[DateTime]::UtcNow.ToString('o')}|
    ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $runRoot "stage-$Stage.json") -Encoding utf8
Set-RefreshPhase 'Complete'
Write-Log 'Read-only stage evaluation completed; no employee values emitted; closed unsaved; workbook hash unchanged. Supervisor save-confirmation rejection is expected.'
