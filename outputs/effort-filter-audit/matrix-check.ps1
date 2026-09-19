$ErrorActionPreference='Stop'
. "$PSScriptRoot/audit.ps1" -FunctionsOnly
$path='CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All.xlsx'
$bytes=ReadBytes (Join-Path $root $path)
$ms=[IO.MemoryStream]::new($bytes,$false)
$z=[IO.Compression.ZipArchive]::new($ms,'Read')
function ReadXml($name){$r=[IO.StreamReader]::new($z.GetEntry($name).Open());try{return [xml]$r.ReadToEnd()}finally{$r.Dispose()}}
try{
 $wb=ReadXml 'xl/workbook.xml';$rels=ReadXml 'xl/_rels/workbook.xml.rels'
 $sheet=$wb.workbook.sheets.sheet | Where-Object name -eq 'AvailabilityDevelopedMAT'
 $rid=$sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
 $target=($rels.Relationships.Relationship | Where-Object Id -eq $rid).Target
 $sheetPath=if($target.StartsWith('/')){$target.TrimStart('/')}else{'xl/'+$target}
 $sx=ReadXml $sheetPath;$ss=ReadXml 'xl/sharedStrings.xml'
 $strings=@($ss.sst.si | ForEach-Object {($_.SelectNodes('.//*[local-name()="t"]') | ForEach-Object InnerText) -join ''})
 $summary=@{};$first1000=@{};$visible=@{};$firstRows=@();$firstOccurrence=@{};$count=0
 foreach($row in $sx.worksheet.sheetData.row){
  if([int]$row.r -le 1){continue};$values=@{}
  foreach($cell in $row.c){$col=[string]$cell.r -replace '\d','';$v=if($cell.t -eq 's'){$strings[[int]$cell.v]}elseif($cell.t -eq 'inlineStr'){$cell.InnerText}else{[string]$cell.v};$values[$col]=$v}
  $role=[string]$values['B'];if($role -notin @('AIN','AINC4','RN')){continue}
  if(!$summary.ContainsKey($role)){$summary[$role]=@{rows=0;positiveC3=0;positiveOriginal=0;positiveAllocation=0;positiveDemand=0};$firstOccurrence[$role]=[int]$row.r}
  $summary[$role].rows++
  foreach($pair in @(@('K','positiveC3'),@('I','positiveOriginal'),@('L','positiveAllocation'),@('H','positiveDemand'))){$n=0.0;if([double]::TryParse([string]$values[$pair[0]],[ref]$n) -and $n -gt 0){$summary[$role][$pair[1]]++}}
  if([int]$row.r -le 1001){if(!$first1000.ContainsKey($role)){$first1000[$role]=0};$first1000[$role]++}
  if($row.hidden -ne '1'){if(!$visible.ContainsKey($role)){$visible[$role]=0};$visible[$role]++}
  if($firstRows.Count -lt 8){$firstRows+=@{row=[int]$row.r;role=$role;resource=$values['E'];facility=$values['A']}}
  $count++
 }
 $m=Extract $bytes
 $q=[regex]::Match($m,'(?s)shared ResRoleAvailabilityDevelopedMATRIX =.*?(?=\r?\nshared )').Value
 $avail=[regex]::Match($m,'(?s)shared #"Availabilities Appended" =.*?(?=\r?\nshared )').Value
 $result=[pscustomobject]@{path=$path;modified=(Get-Item $path).LastWriteTime.ToString('o');summary=$summary;first1000=$first1000;firstOccurrence=$firstOccurrence;visible=$visible;firstRows=$firstRows;matrixQuery=$q;availabilityQuery=$avail}
 $result | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $PSScriptRoot 'matrix-current-check.json')
 $result | ConvertTo-Json -Depth 8
}finally{$z.Dispose();$ms.Dispose()}
