$ErrorActionPreference='Stop'
. "$PSScriptRoot/audit.ps1" -FunctionsOnly
$reports=@()
foreach($unit in @('Unit1','Unit2')){
 $path="CLIENT/DATExx-Whiddon/UNITS/$unit/2. Calculations/Effort.xlsx"
 $b=ReadBytes (Join-Path $root $path);$ms=[IO.MemoryStream]::new($b,$false);$z=[IO.Compression.ZipArchive]::new($ms,'Read')
 function X($name){$e=$z.GetEntry($name);if(!$e){return $null};$r=[IO.StreamReader]::new($e.Open());try{return [xml]$r.ReadToEnd()}finally{$r.Dispose()}}
 try{
  $wb=X 'xl/workbook.xml';$rels=X 'xl/_rels/workbook.xml.rels';$ss=X 'xl/sharedStrings.xml'
  $strings=@($ss.sst.si|ForEach-Object{($_.SelectNodes('.//*[local-name()="t"]')|ForEach-Object InnerText)-join ''})
  foreach($sheet in $wb.workbook.sheets.sheet){
   if($sheet.name -notin @('RoleShiftAvailabilities','ResShiftAllocation','EffortAllMatrixAG1-1D')){continue}
   $rid=$sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships');$target=($rels.Relationships.Relationship|Where-Object Id -eq $rid).Target
   $sp=if($target.StartsWith('/')){$target.TrimStart('/')}else{'xl/'+$target};$sx=X $sp
   $headers=@{};$stats=@{};$resource108=@{};$first1000=@{}
   foreach($row in $sx.worksheet.sheetData.row){
    $values=@{};foreach($cell in $row.c){$col=[string]$cell.r -replace '\d','';$value=if($cell.t -eq 's'){$strings[[int]$cell.v]}elseif($cell.t -eq 'inlineStr'){$cell.InnerText}else{[string]$cell.v};$values[$col]=$value}
    if([int]$row.r -eq 1){foreach($col in $values.Keys){$headers[$col]=$values[$col]};continue}
    $data=@{};foreach($col in $values.Keys){if($headers.ContainsKey($col)){$data[$headers[$col]]=$values[$col]}}
    $role=[string]$data.Role;if(!$stats.ContainsKey($role)){$stats[$role]=@{rows=0;positive=0;resources=@{};missingResource=0;missingDate=0;missingPeriod=0}}
    $stats[$role].rows++;$stats[$role].resources[[string]$data.Resource]=1
    if($data.ContainsKey('Resource') -and [string]::IsNullOrWhiteSpace([string]$data.Resource)){$stats[$role].missingResource++}
    if($data.ContainsKey('Date') -and [string]::IsNullOrWhiteSpace([string]$data.Date)){$stats[$role].missingDate++}
    if($data.ContainsKey('Period') -and [string]::IsNullOrWhiteSpace([string]$data.Period)){$stats[$role].missingPeriod++}
    $val=if($sheet.name -eq 'RoleShiftAvailabilities'){$data.Availability}elseif($sheet.name -eq 'ResShiftAllocation'){$data.ResShiftFTE}else{$data.Demand}
    $n=0.0;if([double]::TryParse([string]$val,[ref]$n)-and $n -gt 0){$stats[$role].positive++}
    if([string]$data.Resource -eq '108'){$key=$role+' / '+[string]$data.Capacity;if(!$resource108.ContainsKey($key)){$resource108[$key]=0};$resource108[$key]++}
    if([int]$row.r -le 1001){if(!$first1000.ContainsKey($role)){$first1000[$role]=0};$first1000[$role]++}
   }
   foreach($role in @($stats.Keys)){$stats[$role].resources=$stats[$role].resources.Count}
   $reports+=[pscustomobject]@{unit=$unit;sheet=$sheet.name;roles=$stats;resource108=$resource108;first1000=$first1000}
  }
 }finally{$z.Dispose();$ms.Dispose()}
}
$reports|ConvertTo-Json -Depth 7|Set-Content (Join-Path $PSScriptRoot 'input-role-check.json')
$reports|ConvertTo-Json -Depth 7
