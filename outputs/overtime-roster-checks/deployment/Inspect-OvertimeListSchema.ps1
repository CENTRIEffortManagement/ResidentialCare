[CmdletBinding()]
param([string] $LogPath='', [string] $StopRequestPath='', [string] $WorkerStatePath='')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
$runFolder=(Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/deployment-current.txt') -Raw).Trim()
$runRoot=Join-Path $repoRoot $runFolder
if(-not $LogPath){$LogPath=Join-Path $runRoot 'list-schema-probe.log'}
if(-not $StopRequestPath){$StopRequestPath=Join-Path $runRoot 'stop.txt'}
$workbookPath=Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx'
function Write-Log([string] $Message,[string] $Level='INFO'){Add-Content -LiteralPath $LogPath -Value "[$([DateTime]::UtcNow.ToString('o'))] [$Level] $Message";Write-Host $Message}
function Write-CurrentStatus([string] $State,[string] $Message=''){}
function Test-StopNowRequested {return (Test-Path -LiteralPath $StopRequestPath) -and ((Get-Content -LiteralPath $StopRequestPath -Raw) -match '(?im)^mode\s*=\s*Now\s*$')}
function Assert-NotStopNowRequested {if(Test-StopNowRequested){throw 'Operator stop observed.'}}
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
if(-not $WorkerStatePath){exit (Invoke-SupervisedExcelRefresh -WorkerScript $PSCommandPath -Parameters @{LogPath=$LogPath;StopRequestPath=$StopRequestPath} -TimeoutSeconds 600)}
$script:WorkerStatePath=$WorkerStatePath
$script:RefreshState=@{phase='Starting';detail=$null;excel=$null;saved=$false;updatedAt=''}
function Release-ProbeCom($Object){if($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)}}
$excel=$null;$book=$null;$queries=$null;$query=$null;$sheet=$null;$table=$null;$qt=$null;$range=$null;$closed=$false;$quit=$false
$before=(Get-FileHash -LiteralPath $workbookPath).Hash
try{
    if(Test-Path -LiteralPath (Join-Path (Split-Path $workbookPath) '~$OvertimeRosterChecks.xlsx')){throw 'Workbook has an Excel lock.'}
    $handle=[IO.File]::Open($workbookPath,'Open','ReadWrite','None');$handle.Dispose()
    $existing=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    Set-RefreshPhase 'CreatingExcel'
    $excel=New-Object -ComObject Excel.Application
    Register-RefreshExcel -Excel $excel -PreExistingExcelIds $existing
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($workbookPath,0,$false)
    Wait-RefreshWorkbookIdentity -Workbook $book -ExpectedPath $workbookPath
    $formula=@'
let
    Navigation = #"IMPORT OvertimeRosterChecks Whiddon Lists",
    Threshold = Navigation{[Id = "ae98d433-faad-4ec0-9209-525ceaa71a29"]}[Items],
    Roles = #"LINK OvertimeRosterChecks Roles",
    Facilities = #"LINK OvertimeRosterChecks Facilities",
    SchemaRows = #table(type table [Source = text, ColumnNames = text], {
        {"Navigation", Text.Combine(Table.ColumnNames(Navigation), " | ")},
        {"Threshold", Text.Combine(Table.ColumnNames(Threshold), " | ")},
        {"Roles", Text.Combine(Table.ColumnNames(Roles), " | ")},
        {"Facilities", Text.Combine(Table.ColumnNames(Facilities), " | ")}
    })
in
    SchemaRows
'@
    $queries=$book.Queries;$query=$queries.Add('OvertimeSchemaProbe',$formula)
    $sheet=$book.Worksheets.Add();$sheet.Name='OvertimeSchemaProbe'
    $table=$sheet.ListObjects.Add(0,'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="OvertimeSchemaProbe";Extended Properties=""',$null,1,$sheet.Range('A1'))
    $qt=$table.QueryTable;$qt.CommandType=2;$qt.CommandText=@('SELECT * FROM [OvertimeSchemaProbe]');$qt.BackgroundQuery=$false
    Set-RefreshPhase 'ReadingListSchemas'
    if(-not $qt.Refresh($false)){throw 'Schema probe did not refresh.'}
    $range=$table.DataBodyRange;$values=$range.Value2
    $rows=@(for($index=1;$index -le $values.GetLength(0);$index++){[ordered]@{Source=[string]$values[$index,1];Columns=[string]$values[$index,2]}})
    $rows | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $runRoot 'live-list-schemas.json') -Encoding utf8
    $rows | Format-Table -AutoSize | Out-String | Write-Host
    Set-RefreshPhase 'ClosingWithoutSaving'
    $book.Close($false);$closed=$true;$excel.Quit();$quit=$true
}finally{
    if($null -ne $book -and -not $closed){$book.Close($false)}
    if($null -ne $excel -and -not $quit){$excel.Quit()}
    foreach($object in @($range,$qt,$table,$sheet,$query,$queries,$book,$excel)){Release-ProbeCom $object}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash -LiteralPath $workbookPath).Hash -ne $before){throw 'Unsaved schema probe changed the workbook.'}
Set-RefreshPhase 'Complete'
Write-Log 'Schema read completed; workbook closed without saving and hash unchanged.'
