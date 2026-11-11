[CmdletBinding()]
param([string]$LogPath='', [string]$StopRequestPath='', [string]$WorkerStatePath='')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
$runRelative='outputs/overtime-roster-checks/clean-base-20261105'
$runRoot=Join-Path $repoRoot $runRelative
$plan=Get-Content -LiteralPath (Join-Path $runRoot 'archive-and-source.json') -Raw|ConvertFrom-Json
if(-not $LogPath){$LogPath=Join-Path $runRoot 'prepare.log'}
if(-not $StopRequestPath){$StopRequestPath=Join-Path $runRoot 'stop.txt'}
function Write-Log([string]$Message,[string]$Level='INFO'){$line="[$([DateTime]::UtcNow.ToString('o'))] [$Level] $Message";Add-Content -LiteralPath $LogPath -Value $line;Write-Host $line}
function Write-CurrentStatus([string]$State,[string]$Message=''){}
function Test-StopNowRequested{(Test-Path -LiteralPath $StopRequestPath) -and ((Get-Content -LiteralPath $StopRequestPath -Raw)-match'(?im)^mode\s*=\s*Now\s*$')}
function Assert-NotStopNowRequested{if(Test-StopNowRequested){throw 'Stop requested.'}}
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
if(-not $WorkerStatePath){exit (Invoke-SupervisedExcelRefresh -WorkerScript $PSCommandPath -Parameters @{LogPath=$LogPath;StopRequestPath=$StopRequestPath} -TimeoutSeconds 180)}
$script:WorkerStatePath=$WorkerStatePath
$script:RefreshState=@{phase='Starting';detail=$null;excel=$null;saved=$false;updatedAt=''}
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'),[ref]$tokens,[ref]$errors)
foreach($name in @('Get-Sha256','Normalize-Formula','Read-MDefinitions','Release-ComObject')){
    $function=@($ast.FindAll({param($node)$node-is[Management.Automation.Language.FunctionDefinitionAst]},$false)|Where-Object Name -eq $name)
    if($function.Count-ne1){throw "Missing approved helper: $name"};. ([scriptblock]::Create($function[0].Extent.Text))
}
$sourcePath=Join-Path $repoRoot $plan.Source
$workbookPath=Join-Path $repoRoot $plan.Workbook
$beforePath=Join-Path $runRoot 'OvertimeRosterChecks-attempt1.m'
$excel=$null;$book=$null;$queries=$null;$pathName=$null;$pathCell=$null;$pathSheet=$null;$lease=$null;$closed=$false;$quit=$false
try{
    Set-RefreshPhase 'Preflight'
    if((Get-Sha256 $sourcePath)-ne$plan.CleanSourceSHA256){throw 'Clean source changed.'}
    if((Get-Sha256 $workbookPath)-ne$plan.ArchivedWorkbookSHA256){throw 'Working workbook changed after archive.'}
    if((Get-Sha256 (Join-Path $repoRoot $plan.ArchivedWorkbook))-ne$plan.ArchivedWorkbookSHA256 -or (Get-Sha256 (Join-Path $repoRoot $plan.ArchivedSource))-ne$plan.ArchivedSourceSHA256){throw 'Attempt1 archive changed.'}
    $before=Read-MDefinitions $beforePath;$clean=Read-MDefinitions $sourcePath
    if($clean.Count-ne10){throw 'Expected ten import/path definitions.'}
    foreach($name in $clean.Keys){if(-not$before.Contains($name)-or(Normalize-Formula $clean[$name])-cne(Normalize-Formula $before[$name])){throw "Retained import/path changed: $name"}}
    . (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/RefreshRunGate.ps1')
    $lease=Enter-RefreshRunGate -DateRoot (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon') -Exclusive
    if(Test-Path -LiteralPath (Join-Path (Split-Path $workbookPath) '~$OvertimeRosterChecks.xlsx')){throw 'Workbook is open.'}
    $handle=[IO.File]::Open($workbookPath,'Open','ReadWrite','None');$handle.Dispose()
    $snapshot=New-RefreshFileSnapshot -Path $workbookPath
    $existing=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
    Set-RefreshPhase 'CreatingExcel'
    $excel=New-Object -ComObject Excel.Application;Register-RefreshExcel -Excel $excel -PreExistingExcelIds $existing
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($workbookPath,0,$false);Wait-RefreshWorkbookIdentity -Workbook $book -ExpectedPath $workbookPath
    $queries=$book.Queries
    foreach($name in $before.Keys){$q=$null;try{$q=$queries.Item($name);if((Normalize-Formula ([string]$q.Formula))-cne(Normalize-Formula $before[$name])){throw "Installed baseline differs: $name"}}finally{Release-ComObject $q}}
    $pathName=$book.Names.Item('FilePathUrl');$pathCell=$pathName.RefersToRange;$pathSheet=$pathCell.Worksheet
    if([int]$pathCell.Cells.Count-ne1){throw 'Expected one saved-path cell.'}
    $keepSheet=[string]$pathSheet.Name
    Set-RefreshPhase 'RemovingCalculationDefinitions'
    $retired=@($before.Keys|Where-Object{-not$clean.Contains($_)})
    [array]::Reverse($retired)
    foreach($name in $retired){$q=$null;try{$q=$queries.Item($name);$q.Delete()}finally{Release-ComObject $q}}
    Set-RefreshPhase 'RemovingCalculationSheetsAndConnections'
    for($index=[int]$book.Worksheets.Count;$index-ge1;$index--){$sheet=$null;try{$sheet=$book.Worksheets.Item($index);if([string]$sheet.Name-cne$keepSheet){$sheet.Delete()}else{$sheet=$null}}finally{Release-ComObject $sheet}}
    Set-RefreshPhase 'RemovingOutputConnections'
    for($index=[int]$book.Connections.Count;$index-ge1;$index--){$connection=$null;try{$connection=$book.Connections.Item($index);$connection.Delete()}finally{Release-ComObject $connection}}
    Set-RefreshPhase 'ClearingPathSheet'
    for($index=[int]$pathSheet.ListObjects.Count;$index-ge1;$index--){$table=$null;try{$table=$pathSheet.ListObjects.Item($index);$table.Delete()}finally{Release-ComObject $table}}
    # All other saved names were local to the deleted sheets; retain FilePathUrl.
    $used=$null;try{$used=$pathSheet.UsedRange;$used.Clear()|Out-Null}finally{Release-ComObject $used}
    $pathCell=$pathName.RefersToRange
    $pathSheet.Range('A1').Value2='FilePath'
    $pathCell.Formula='=CELL("filename",A1)';$pathCell.Calculate()
    if([string]$pathCell.Value2-notmatch'\[OvertimeRosterChecks\.xlsx\]'){throw 'FilePathUrl did not recalculate.'}
    foreach($name in $clean.Keys){$q=$null;try{$q=$queries.Item($name);if((Normalize-Formula ([string]$q.Formula))-cne(Normalize-Formula $clean[$name])){throw "Retained definition differs: $name"}}finally{Release-ComObject $q}}
    if([int]$book.Worksheets.Count-ne1-or[int]$book.Connections.Count-ne0){throw 'Unexpected sheets or connections remain.'}
    Assert-NotStopNowRequested;Assert-RefreshFileSnapshot $snapshot
    if((Get-Sha256 $sourcePath)-ne$plan.CleanSourceSHA256){throw 'Source changed during preparation.'}
    Set-RefreshPhase 'Saving';$book.Save();if(-not$book.Saved){throw 'Save was not confirmed.'};$script:RefreshState.saved=$true
    Set-RefreshPhase 'Closing';$book.Close($false);$closed=$true
    Set-RefreshPhase 'Quitting';$excel.Quit();$quit=$true
    foreach($object in @($pathCell,$pathName,$pathSheet,$queries,$book,$excel)){Release-ComObject $object}
    $pathCell=$null;$pathName=$null;$pathSheet=$null;$queries=$null;$book=$null;$excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $afterPath=Join-Path $runRoot 'OvertimeRosterChecks.m'
    & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') -WorkbookPath $workbookPath -OutputMPath $afterPath -EvidencePath (Join-Path $runRoot 'clean-base-extraction.json')|Out-Null
    $actual=Read-MDefinitions $afterPath
    if($actual.Count-ne10){throw 'Re-extraction did not contain exactly ten definitions.'}
    foreach($name in $clean.Keys){if(-not$actual.Contains($name)-or(Normalize-Formula $actual[$name])-cne(Normalize-Formula $clean[$name])){throw "Re-extracted definition differs: $name"}}
    if((Get-Sha256 (Join-Path $repoRoot $plan.ArchivedWorkbook))-ne$plan.ArchivedWorkbookSHA256-or(Get-Sha256 (Join-Path $repoRoot $plan.ArchivedSource))-ne$plan.ArchivedSourceSHA256){throw 'Archive changed during preparation.'}
    [ordered]@{Status='Prepared';Source=$plan.Source;Workbook=$plan.Workbook;SourceSHA256=$plan.CleanSourceSHA256;WorkbookSHA256=(Get-Sha256 $workbookPath);QueryCount=10;Queries=@($clean.Keys);Worksheet=$keepSheet;Connections=0;ReExtractionMatches=$true;ExtractedM=($runRelative+'/OvertimeRosterChecks.m');ArchivedSource=$plan.ArchivedSource;ArchivedWorkbook=$plan.ArchivedWorkbook;ArchivesUnchanged=$true;Saved=$true;Closed=$true;Refreshed=$false;BDWorkbooksChangedDuringReset=$false;Donor=$plan.Donor;RecordedAtUTC=[DateTime]::UtcNow.ToString('o')}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $runRoot 'clean-base-receipt.json') -Encoding utf8
    Set-RefreshPhase 'Complete';Write-Log 'Clean imports/path base saved, closed and re-extracted; no refresh performed.'
}catch{$script:RefreshState.error=@{Phase=$script:RefreshState.phase;Message=$_.Exception.Message;Position=$_.InvocationInfo.PositionMessage;Stack=$_.ScriptStackTrace};Write-RefreshWorkerState;Write-Log ($_.Exception.Message+' '+$_.InvocationInfo.PositionMessage) 'ERROR';throw}
finally{
    if($null-ne$book-and-not$closed){try{$book.Close($false)}catch{}}
    if($null-ne$excel-and-not$quit){try{$excel.Quit()}catch{}}
    foreach($object in @($pathCell,$pathName,$pathSheet,$queries,$book,$excel)){Release-ComObject $object}
    if($null-ne$lease){$lease.Dispose()};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
