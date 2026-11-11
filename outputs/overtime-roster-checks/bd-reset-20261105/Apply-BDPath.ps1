[CmdletBinding()]
param([string]$LogPath='',[string]$StopRequestPath='',[string]$WorkerStatePath='')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
$runRelative='outputs/overtime-roster-checks/bd-reset-20261105'
$runRoot=Join-Path $repoRoot $runRelative
$plan=Get-Content -LiteralPath (Join-Path $runRoot 'move-and-path-plan.json') -Raw|ConvertFrom-Json
if(-not$LogPath){$LogPath=Join-Path $runRoot 'path-update.log'}
if(-not$StopRequestPath){$StopRequestPath=Join-Path $runRoot 'stop.txt'}
function Write-Log([string]$Message,[string]$Level='INFO'){$line="[$([DateTime]::UtcNow.ToString('o'))] [$Level] $Message";Add-Content -LiteralPath $LogPath -Value $line;Write-Host $line}
function Write-CurrentStatus([string]$State,[string]$Message=''){}
function Test-StopNowRequested{(Test-Path -LiteralPath $StopRequestPath)-and((Get-Content -LiteralPath $StopRequestPath -Raw)-match'(?im)^mode\s*=\s*Now\s*$')}
function Assert-NotStopNowRequested{if(Test-StopNowRequested){throw 'Stop requested.'}}
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
if(-not$WorkerStatePath){exit(Invoke-SupervisedExcelRefresh -WorkerScript $PSCommandPath -Parameters @{LogPath=$LogPath;StopRequestPath=$StopRequestPath} -TimeoutSeconds 180)}
$script:WorkerStatePath=$WorkerStatePath
$script:RefreshState=@{phase='Starting';detail=$null;excel=$null;saved=$false;updatedAt=''}
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'),[ref]$tokens,[ref]$errors)
foreach($name in @('Get-Sha256','Normalize-Formula','Read-MDefinitions','Release-ComObject')){$fn=@($ast.FindAll({param($n)$n-is[Management.Automation.Language.FunctionDefinitionAst]},$false)|Where-Object Name -eq $name);if($fn.Count-ne1){throw 'Missing approved sync helper.'};.([scriptblock]::Create($fn[0].Extent.Text))}
$bookPath=Join-Path $repoRoot $plan.Workbook
$sourcePath=Join-Path $repoRoot $plan.Source
$excel=$null;$book=$null;$queries=$null;$pathName=$null;$pathCell=$null;$lease=$null;$closed=$false;$quit=$false
try{
    Set-RefreshPhase 'Preflight'
    if((Get-Sha256 $bookPath)-ne$plan.WorkbookSHA256Before-or(Get-Sha256 $sourcePath)-ne$plan.SourceSHA256){throw 'Workbook/source changed after path preparation.'}
    $before=Read-MDefinitions (Join-Path $repoRoot $plan.BaselineM)
    $approved=Read-MDefinitions $sourcePath
    if($before.Count-ne10-or$approved.Count-ne10){throw 'Expected ten source definitions.'}
    . (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/RefreshRunGate.ps1')
    $lease=Enter-RefreshRunGate -DateRoot (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon') -Exclusive
    if(Test-Path -LiteralPath (Join-Path (Split-Path $bookPath) '~$OvertimeRosterChecks.xlsx')){throw 'Workbook is open.'}
    $handle=[IO.File]::Open($bookPath,'Open','ReadWrite','None');$handle.Dispose()
    $snapshot=New-RefreshFileSnapshot -Path $bookPath
    $existing=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
    Set-RefreshPhase 'CreatingExcel'
    $excel=New-Object -ComObject Excel.Application;Register-RefreshExcel -Excel $excel -PreExistingExcelIds $existing
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($bookPath,0,$false);Wait-RefreshWorkbookIdentity -Workbook $book -ExpectedPath $bookPath
    $queries=$book.Queries
    foreach($name in $before.Keys){$q=$null;try{$q=$queries.Item($name);if((Normalize-Formula([string]$q.Formula))-cne(Normalize-Formula $before[$name])){throw "Installed baseline differs: $name"}}finally{Release-ComObject $q}}
    Set-RefreshPhase 'UpdatingPathDefinitions'
    foreach($name in $approved.Keys){$q=$null;try{if($before.Contains($name)){$q=$queries.Item($name);$q.Formula=[string]$approved[$name]}else{$q=$queries.Add($name,[string]$approved[$name])};if((Normalize-Formula([string]$q.Formula))-cne(Normalize-Formula $approved[$name])){throw "Updated definition differs: $name"}}finally{Release-ComObject $q}}
    foreach($name in $before.Keys){if(-not$approved.Contains($name)){$q=$null;try{$q=$queries.Item($name);$q.Delete()}finally{Release-ComObject $q}}}
    Set-RefreshPhase 'RecalculatingSavedPath'
    $pathName=$book.Names.Item('FilePathUrl');$pathCell=$pathName.RefersToRange
    if([int]$pathCell.Cells.Count-ne1){throw 'Expected one FilePathUrl cell.'}
    $pathCell.Formula='=CELL("filename",A1)';$pathCell.Calculate()
    if([string]$pathCell.Value2-notmatch'(?i)\\UNITS\\BD\\2\. Calculations\\\[OvertimeRosterChecks\.xlsx\]'){throw 'FilePathUrl does not identify BD/2. Calculations.'}
    $savedPath=[string]$pathCell.Value2
    Assert-NotStopNowRequested;Assert-RefreshFileSnapshot $snapshot
    if((Get-Sha256 $sourcePath)-ne$plan.SourceSHA256){throw 'Approved source changed.'}
    Set-RefreshPhase 'Saving';$book.Save();if(-not$book.Saved){throw 'Save not confirmed.'};$script:RefreshState.saved=$true
    Set-RefreshPhase 'Closing';$book.Close($false);$closed=$true
    Set-RefreshPhase 'Quitting';$excel.Quit();$quit=$true
    foreach($object in @($pathCell,$pathName,$queries,$book,$excel)){Release-ComObject $object}
    $pathCell=$null;$pathName=$null;$queries=$null;$book=$null;$excel=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()
    & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') -WorkbookPath $bookPath -OutputMPath (Join-Path $runRoot 'OvertimeRosterChecks.m') -EvidencePath (Join-Path $runRoot 'Overtime-after-path-extraction.json')|Out-Null
    $actual=Read-MDefinitions (Join-Path $runRoot 'OvertimeRosterChecks.m')
    if($actual.Count-ne$approved.Count){throw 'Re-extracted query set differs.'}
    foreach($name in $approved.Keys){if(-not$actual.Contains($name)-or(Normalize-Formula $actual[$name])-cne(Normalize-Formula $approved[$name])){throw "Re-extracted definition differs: $name"}}
    [ordered]@{Status='Moved and path updated';Workbook=$plan.Workbook;Source=$plan.Source;SourceSHA256=$plan.SourceSHA256;WorkbookSHA256=(Get-Sha256 $bookPath);QueryCount=$actual.Count;Queries=@($actual.Keys);ReExtractionMatches=$true;FilePathUrl=$savedPath;Saved=$true;Closed=$true;Refreshed=$false;ExtractedM=($runRelative+'/OvertimeRosterChecks.m');RecordedAtUTC=[DateTime]::UtcNow.ToString('o')}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $runRoot 'Overtime-path-receipt.json') -Encoding utf8
    Set-RefreshPhase 'Complete';Write-Log 'Moved workbook path and source verified; saved and closed without data refresh.'
}catch{$script:RefreshState.error=@{Phase=$script:RefreshState.phase;Message=$_.Exception.Message;Position=$_.InvocationInfo.PositionMessage};Write-RefreshWorkerState;Write-Log ($_.Exception.Message+' '+$_.InvocationInfo.PositionMessage) 'ERROR';throw}
finally{
    if($null-ne$book-and-not$closed){try{$book.Close($false)}catch{}}
    if($null-ne$excel-and-not$quit){try{$excel.Quit()}catch{}}
    foreach($object in @($pathCell,$pathName,$queries,$book,$excel)){Release-ComObject $object}
    if($null-ne$lease){$lease.Dispose()};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
