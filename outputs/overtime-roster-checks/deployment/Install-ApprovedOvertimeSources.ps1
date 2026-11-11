[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateSet('Core','Shifts')] [string] $Target,
    [string] $RunFolder = '',
    [string] $LogPath = '',
    [string] $StopRequestPath = '',
    [string] $WorkerStatePath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))
if (-not $RunFolder) { $RunFolder = (Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/deployment-current.txt') -Raw).Trim() }
$runRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot $RunFolder))
if (-not $runRoot.StartsWith(($repoRoot.TrimEnd('\') + '\'), [StringComparison]::OrdinalIgnoreCase)) { throw 'Deployment evidence must remain within the repository.' }
if (-not $LogPath) { $LogPath = Join-Path $runRoot "$Target-install.log" }
if (-not $StopRequestPath) { $StopRequestPath = Join-Path $runRoot 'stop.txt' }
$targets = @{
    Core = @{Workbook='CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx';Source='Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq';Outputs=@(
        @{Query='OvertimeRosterChecks_ShiftResults';Sheet='ShiftResults'},
        @{Query='OvertimeRosterChecks_FortnightResults';Sheet='FortnightResults'},
        @{Query='OvertimeRosterChecks_ExcessSegments';Sheet='ExcessSegments'})}
    Shifts = @{Workbook='CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shifts.xlsx';Source='Workflows/ResidentialCare/OvertimeRosterChecks/BD/Shifts.pq';Outputs=@(
        @{Query='OvertimeRosterChecks';Sheet='OvertimeRosterChecks'})}
}
$selection = $targets[$Target]
$workbookPath = Join-Path $repoRoot $selection.Workbook
$sourcePath = Join-Path $repoRoot $selection.Source
$beforePath = Join-Path $runRoot "$Target-before.m"
$beforeEvidence = Get-Content -LiteralPath (Join-Path $runRoot "$Target-before-extraction.json") -Raw | ConvertFrom-Json
$approvedReceipt = Get-Content -LiteralPath (Join-Path $repoRoot 'outputs/overtime-roster-checks/implementation-verification.json') -Raw | ConvertFrom-Json
$expectedSourceHash = if ($Target -eq 'Core') { $approvedReceipt.CoreSHA256 } else { $approvedReceipt.ShiftsSHA256 }

# Reuse the approved Excel-object-model sync mechanism; never rewrite an XLSX package.
$parseTokens=$null; $parseErrors=$null
$syncAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'),[ref]$parseTokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw 'Approved synchronization mechanism has syntax errors.' }
foreach ($name in @('Get-Sha256','Normalize-Formula','Read-MDefinitions','Release-ComObject')) {
    $function=@($syncAst.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object Name -eq $name)
    if ($function.Count -ne 1) { throw "Missing approved sync helper $name" }
    . ([scriptblock]::Create($function[0].Extent.Text))
}
function Write-Log([string] $Message, [string] $Level='INFO') {
    $line="[$([DateTime]::UtcNow.ToString('o'))] [$Level] $Target $Message"
    Write-Host $line
    Add-Content -LiteralPath (Join-Path $runRoot "$Target-install.log") -Value $line
}
function Write-CurrentStatus([string] $State,[string] $Message='') { }
function Test-StopNowRequested {
    return (Test-Path -LiteralPath $StopRequestPath) -and ((Get-Content -LiteralPath $StopRequestPath -Raw) -match '(?im)^mode\s*=\s*Now\s*$')
}
function Assert-NotStopNowRequested { if(Test-StopNowRequested){throw 'Operator stop-now request observed.'} }
$WorkerStatePath = if ($WorkerStatePath) { [IO.Path]::GetFullPath($WorkerStatePath) } else { '' }
. (Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/ExcelRefreshSafety.ps1')
if (-not $WorkerStatePath) {
    $arguments=@{Target=$Target;RunFolder=$RunFolder;LogPath=$LogPath;StopRequestPath=$StopRequestPath}
    exit (Invoke-SupervisedExcelRefresh -WorkerScript $PSCommandPath -Parameters $arguments -TimeoutSeconds 1800)
}
$script:WorkerStatePath=$WorkerStatePath
$script:RefreshState=@{phase='Starting';detail=$null;excel=$null;saved=$false;updatedAt=''}

function Assert-Closed([string] $Path) {
    if (Test-Path -LiteralPath (Join-Path (Split-Path $Path -Parent) ('~$'+[IO.Path]::GetFileName($Path)))) { throw "Excel lock file exists: $Path" }
    $handle=[IO.File]::Open($Path,'Open','ReadWrite','None')
    $handle.Dispose()
}
function Get-XmlPart($Archive,[string] $Name) {
    $part=$Archive.GetEntry($Name)
    if ($null -eq $part) { return $null }
    $reader=[IO.StreamReader]::new($part.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Resolve-Part([string] $Base,[string] $Relative) {
    return ([uri]::new([uri]('http://package/'+$Base),$Relative)).AbsolutePath.TrimStart('/')
}
function Read-PackageLayout([string] $Path) {
    $archive=[IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $book=Get-XmlPart $archive 'xl/workbook.xml'
        $rels=Get-XmlPart $archive 'xl/_rels/workbook.xml.rels'
        $sheetRows=@(); $tables=@()
        foreach ($sheet in $book.SelectNodes('//*[local-name()="sheet"]')) {
            $relationshipId=$sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $relationship=@($rels.DocumentElement.ChildNodes | Where-Object Id -eq $relationshipId)
            if ($relationship.Count -ne 1) { throw 'Ambiguous worksheet package relationship.' }
            $part=Resolve-Part 'xl/workbook.xml' $relationship[0].Target
            $sheetRows += [string]$sheet.name
            $sheetXml=Get-XmlPart $archive $part
            $relPart=([IO.Path]::GetDirectoryName($part).Replace('\','/'))+'/_rels/'+[IO.Path]::GetFileName($part)+'.rels'
            $sheetRels=Get-XmlPart $archive $relPart
            foreach ($tablePart in $sheetXml.SelectNodes('//*[local-name()="tablePart"]')) {
                $id=$tablePart.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                $tableRel=@($sheetRels.DocumentElement.ChildNodes | Where-Object Id -eq $id)
                $tableXml=Get-XmlPart $archive (Resolve-Part $part $tableRel[0].Target)
                $tables += [pscustomobject]@{Name=[string]$tableXml.DocumentElement.name;Sheet=[string]$sheet.name;QueryTable=($tableXml.DocumentElement.GetAttribute('tableType') -eq 'queryTable')}
            }
        }
        $connectionXml=Get-XmlPart $archive 'xl/connections.xml'
        $connections=@(if($null -ne $connectionXml){foreach($connection in $connectionXml.DocumentElement.ChildNodes){
            $db=$connection.SelectSingleNode('./*[local-name()="dbPr"]')
            [pscustomobject]@{Name=[string]$connection.name;Mashup=($null -ne $db -and [string]$db.connection -match 'Microsoft\.Mashup\.OleDb\.1')}
        }})
        return @{Sheets=$sheetRows;Tables=$tables;Connections=$connections}
    } finally { $archive.Dispose() }
}
function Add-LoadedOutput($Book,[string] $Query,[string] $Sheet) {
    $worksheet=$null; $destination=$null; $table=$null; $queryTable=$null; $connection=$null
    try {
        $worksheet=$Book.Worksheets.Add()
        $worksheet.Name=$Sheet
        $destination=$worksheet.Range('A1')
        $connectionString='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="'+$Query.Replace('"','""')+'";Extended Properties=""'
        $table=$worksheet.ListObjects.Add(0,$connectionString,$null,1,$destination)
        $table.Name=$Query
        $queryTable=$table.QueryTable
        $queryTable.CommandType=2
        $queryTable.CommandText=@('SELECT * FROM ['+$Query.Replace(']',']]')+']')
        $queryTable.BackgroundQuery=$false
        $queryTable.RefreshOnFileOpen=$false
        $queryTable.EnableRefresh=$true
        $connection=$queryTable.WorkbookConnection
        $connection.RefreshWithRefreshAll=$true
        if ([string]$table.Name -cne $Query) { throw "Incorrect output table name for $Query" }
        Write-Log "Installed output table $Query on $Sheet; refresh deferred to the hardened worker."
    } finally {
        foreach($object in @($connection,$queryTable,$table,$destination,$worksheet)){Release-ComObject $object}
    }
}

$excel=$null; $book=$null; $queries=$null; $ownsExcel=$false; $closed=$false; $quit=$false
try {
    Set-RefreshPhase 'Preflight'
    Assert-NotStopNowRequested
    Assert-Closed $workbookPath
    if ((Get-Sha256 $sourcePath) -ne $expectedSourceHash) { throw 'Approved M source changed after native validation.' }
    if ((Get-Sha256 $workbookPath) -ne $beforeEvidence.workbookSHA256) { throw 'Workbook changed after fresh extraction.' }
    $before=Read-MDefinitions $beforePath
    $approved=Read-MDefinitions $sourcePath
    if ($Target -eq 'Shifts' -and @($before.Keys | Where-Object {-not $approved.Contains($_)}).Count) { throw 'BD synchronization would remove an existing query.' }
    $layout=Read-PackageLayout $workbookPath
    foreach($output in $selection.Outputs){if($layout.Sheets -contains $output.Sheet -or $layout.Tables.Name -contains $output.Query){throw 'New output destination already exists; reconcile instead of overwriting.'}}
    $backupDirectory=Join-Path $runRoot 'backups'
    [void][IO.Directory]::CreateDirectory($backupDirectory)
    $backup=Join-Path $backupDirectory "$Target.before.xlsx"
    if(Test-Path -LiteralPath $backup){if((Get-Sha256 $backup) -ne $beforeEvidence.workbookSHA256){throw 'Backup differs from preflight workbook.'}}
    else {Copy-Item -LiteralPath $workbookPath -Destination $backup}
    if((Get-Sha256 $backup) -ne $beforeEvidence.workbookSHA256){throw 'Backup verification failed.'}
    $preExisting=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    Assert-Closed $workbookPath
    $fileSnapshot=New-RefreshFileSnapshot -Path $workbookPath
    Set-RefreshPhase 'CreatingExcel'
    $excel=New-Object -ComObject Excel.Application
    Register-RefreshExcel -Excel $excel -PreExistingExcelIds $preExisting
    $ownsExcel=$true
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AskToUpdateLinks=$false; $excel.AutomationSecurity=3
    Set-RefreshPhase 'Opening'
    $book=$excel.Workbooks.Open($workbookPath,0,$false)
    Wait-RefreshWorkbookIdentity -Workbook $book -ExpectedPath $workbookPath
    $queries=$book.Queries
    foreach($name in $before.Keys){
        $query=$null
        try{$query=$queries.Item($name); if([string]$query.Name -cne $name -or (Normalize-Formula ([string]$query.Formula)) -cne (Normalize-Formula $before[$name])){throw "Native definition differs from extraction: $name"}}
        finally{Release-ComObject $query}
    }
    Set-RefreshPhase 'InstallingQueries'
    $removed=@(); $updated=@(); $added=@()
    if($Target -eq 'Core'){
        # Retire only copied donor query loads/connections; keep static tables, names and path input.
        foreach($item in $layout.Tables | Where-Object QueryTable){
            $sheet=$null; $table=$null
            try{$sheet=$book.Worksheets.Item($item.Sheet);$table=$sheet.ListObjects.Item($item.Name);$table.Delete()}
            finally{Release-ComObject $table;Release-ComObject $sheet}
        }
        foreach($item in $layout.Connections | Where-Object Mashup){$connection=$null;try{$connection=$book.Connections.Item($item.Name);$connection.Delete()}finally{Release-ComObject $connection}}
        foreach($name in $before.Keys){if(-not $approved.Contains($name)){$query=$null;try{$query=$queries.Item($name);$query.Delete();$removed+=$name}finally{Release-ComObject $query}}}
    }
    foreach($name in $approved.Keys){
        Assert-NotStopNowRequested
        $query=$null
        try{
            if($before.Contains($name)){$query=$queries.Item($name);if((Normalize-Formula ([string]$query.Formula)) -cne (Normalize-Formula $approved[$name])){$query.Formula=[string]$approved[$name];$updated+=$name}}
            else{$created=$queries.Add($name,[string]$approved[$name]);Release-ComObject $created;$query=$queries.Item($name);$added+=$name}
            if((Normalize-Formula ([string]$query.Formula)) -cne (Normalize-Formula $approved[$name])){throw "Installed query differs: $name"}
        }finally{Release-ComObject $query}
    }
    if($Target -eq 'Core'){
        Set-RefreshPhase 'UpdatingCopiedPathInput'
        $pathName=$null; $pathRange=$null
        try{
            $pathName=$book.Names.Item('FilePathUrl');$pathRange=$pathName.RefersToRange
            if([int]$pathRange.Cells.Count -ne 1){throw 'Copied FilePathUrl must be a single named cell.'}
            $pathRange.Formula='=CELL("filename",A1)'
            $pathRange.Calculate()
            if([string]$pathRange.Value2 -notmatch '\[OvertimeRosterChecks\.xlsx\]'){throw 'Copied FilePathUrl did not recalculate for its new filename.'}
        }finally{Release-ComObject $pathRange;Release-ComObject $pathName}
    }
    Set-RefreshPhase 'InstallingOutputLoads'
    foreach($output in $selection.Outputs){Add-LoadedOutput $book $output.Query $output.Sheet}
    foreach($name in $approved.Keys){$query=$null;try{$query=$queries.Item($name);if((Normalize-Formula ([string]$query.Formula)) -cne (Normalize-Formula $approved[$name])){throw "Pre-save query differs: $name"}}finally{Release-ComObject $query}}
    if((Get-Sha256 $sourcePath) -ne $expectedSourceHash){throw 'Approved source changed during installation.'}
    # Excel owns the target's write handle here. Use the runner's metadata snapshot,
    # then hash the saved file only after Excel closes it, as in the approved sync mechanism.
    Assert-RefreshFileSnapshot $fileSnapshot
    Set-RefreshPhase 'Saving'
    Assert-NotStopNowRequested
    $book.Save()
    if(-not $book.Saved){throw 'Excel did not confirm the synchronized workbook save.'}
    $script:RefreshState.saved=$true
    Set-RefreshPhase 'Closing'
    $book.Close($false);$closed=$true
    Set-RefreshPhase 'Quitting'
    $excel.Quit();$quit=$true
    Release-ComObject $queries;$queries=$null;Release-ComObject $book;$book=$null;Release-ComObject $excel;$excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    Assert-Closed $workbookPath
    Set-RefreshPhase 'ReExtracting'
    $afterPath=Join-Path $runRoot "$Target-after-sync.m"
    & (Join-Path $repoRoot 'outputs/te-availability-role-fix/extract-powerquery.ps1') -WorkbookPath $workbookPath -OutputMPath $afterPath -EvidencePath (Join-Path $runRoot "$Target-after-sync-extraction.json") | Out-Null
    $after=Read-MDefinitions $afterPath
    if($after.Count -ne $approved.Count -or @($after.Keys | Where-Object {-not $approved.Contains($_)}).Count){throw 'Re-extracted query names differ from the approved source.'}
    foreach($name in $approved.Keys){if((Normalize-Formula $after[$name]) -cne (Normalize-Formula $approved[$name])){throw "Re-extracted formula differs: $name"}}
    [ordered]@{Target=$Target;Workbook=$selection.Workbook;Source=$selection.Source;SourceSHA256=$expectedSourceHash;WorkbookSHA256Before=$beforeEvidence.workbookSHA256;WorkbookSHA256After=(Get-Sha256 $workbookPath);RemovedQueries=$removed;UpdatedQueries=$updated;AddedQueries=$added;VerifiedQueryCount=$after.Count;LoadedOutputs=@($selection.Outputs.Query);ReExtractionMatches=$true;Saved=$true;Closed=$true;Refreshed=$false} |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $runRoot "$Target-sync-receipt.json") -Encoding utf8
    Set-RefreshPhase 'Complete'
    Write-Log 'Synchronization, save, close and exact re-extraction verification passed.'
    exit 0
} catch {
    $script:RefreshState.error=@{Phase=$script:RefreshState.phase;Message=$_.Exception.Message}
    Write-RefreshWorkerState
    Write-Log $_.Exception.Message 'ERROR'
    throw
} finally {
    if($null -ne $book -and -not $closed){try{$book.Close($false)}catch{Write-Log $_.Exception.Message 'WARN'}}
    if($null -ne $excel -and $ownsExcel -and -not $quit){try{$excel.Quit()}catch{Write-Log $_.Exception.Message 'WARN'}}
    foreach($object in @($queries,$book,$excel)){Release-ComObject $object}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
