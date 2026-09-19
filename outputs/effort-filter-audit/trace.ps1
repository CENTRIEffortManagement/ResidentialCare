param([string[]]$Paths)
. "$PSScriptRoot/audit.ps1" -FunctionsOnly
function XmlEntry($z,$name){$e=$z.GetEntry($name);if(!$e){return $null};$r=[IO.StreamReader]::new($e.Open());try{return [xml]$r.ReadToEnd()}finally{$r.Dispose()}}
function ColNumber($s){$n=0;foreach($c in $s.ToCharArray()){$n=$n*26+([int]$c-64)};return $n}
foreach($path in $Paths){
 $b=ReadBytes (Join-Path $root $path)
 $m=Extract $b
 $name=($path -replace '[\\/: ]','_')
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "$name.m"),$m)
 $ms=[IO.MemoryStream]::new($b,$false);$z=[IO.Compression.ZipArchive]::new($ms,'Read')
 try {
 $strings=[Collections.Generic.List[string]]::new();$ss=XmlEntry $z 'xl/sharedStrings.xml'
 if($ss){foreach($si in $ss.sst.si){$strings.Add(($si.SelectNodes('.//*[local-name()="t"]') | ForEach-Object {$_.InnerText}) -join '')}}
 $wb=XmlEntry $z 'xl/workbook.xml';$wr=XmlEntry $z 'xl/_rels/workbook.xml.rels'
 $out=@()
 foreach($sheet in $wb.workbook.sheets.sheet){
 $rid=$sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
 $target=($wr.Relationships.Relationship | Where-Object Id -eq $rid).Target
 $sp=if($target.StartsWith('/')){$target.TrimStart('/')}else{'xl/'+$target}
 $sx=XmlEntry $z $sp
 $relsPath=($sp -replace '([^/]+)$','_rels/$1.rels');$sr=XmlEntry $z $relsPath
 if(!$sr){continue}
 foreach($rel in $sr.Relationships.Relationship){
 if($rel.Type -notmatch '/table$'){continue}
 $uri=[uri]::new([uri]('http://local/'+$sp),[string]$rel.Target);$tx=XmlEntry $z $uri.AbsolutePath.TrimStart('/')
 $table=$tx.table;$cols=@($table.tableColumns.tableColumn | ForEach-Object {$_.name})
 $wanted=@($cols | Where-Object {$_ -match '^(Role|RolesList|RoleList|PreferredRole|Preferred Role|Facility|Facility1|Facility2|DateAlignment|EffortType|AvailabilityType|Attribute|Value|Column1)$'})
 if(!$wanted.Count){continue}
 $bounds=[regex]::Match($table.ref,'^([A-Z]+)(\d+):([A-Z]+)(\d+)$');if(!$bounds.Success){continue}
 $startCol=ColNumber $bounds.Groups[1].Value;$first=[int]$bounds.Groups[2].Value;$last=[int]$bounds.Groups[4].Value
 $maps=@{};foreach($col in $wanted){$maps[$col]=@{}}
 foreach($row in $sx.worksheet.sheetData.row){if([int]$row.r -le $first -or [int]$row.r -gt $last){continue}
 foreach($cell in $row.c){$cn=ColNumber ([string]$cell.r -replace '\d','');$index=$cn-$startCol;if($index -lt 0 -or $index -ge $cols.Count){continue};$col=$cols[$index];if(!$maps.ContainsKey($col)){continue}
 $value=if($cell.t -eq 's'){$strings[[int]$cell.v]}elseif($cell.t -eq 'inlineStr'){$cell.InnerText}else{[string]$cell.v}
 if(!$maps[$col].ContainsKey($value)){$maps[$col][$value]=0};$maps[$col][$value]++
 }}
 $out+= [pscustomobject]@{table=$table.name;sheet=$sheet.name;ref=$table.ref;columns=$cols;counts=$maps;filter=$table.autoFilter.OuterXml;sheetFilter=$sx.worksheet.autoFilter.OuterXml}
 }
 }
 $result=[pscustomobject]@{path=$path;modified=(Get-Item $path).LastWriteTime.ToString('o');tables=$out}
 $result | ConvertTo-Json -Depth 9 | Set-Content (Join-Path $PSScriptRoot "$name.tables.json")
 $result | ConvertTo-Json -Depth 9
 }finally{$z.Dispose();$ms.Dispose()}
}
