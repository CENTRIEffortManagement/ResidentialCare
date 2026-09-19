param([switch]$FunctionsOnly)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
function ReadBytes($p) {
 $s=[IO.File]::Open($p,'Open','Read','ReadWrite'); $m=[IO.MemoryStream]::new()
 try {$s.CopyTo($m); return ,$m.ToArray()} finally {$s.Dispose();$m.Dispose()}
}
function GitBytes($spec) {
 $p=[Diagnostics.Process]::new();$p.StartInfo.FileName='git';$p.StartInfo.WorkingDirectory=$root
 $p.StartInfo.ArgumentList.Add('cat-file');$p.StartInfo.ArgumentList.Add('blob');$p.StartInfo.ArgumentList.Add($spec)
 $p.StartInfo.UseShellExecute=$false;$p.StartInfo.RedirectStandardOutput=$true;$p.StartInfo.CreateNoWindow=$true
 $null=$p.Start();$m=[IO.MemoryStream]::new();$p.StandardOutput.BaseStream.CopyTo($m);$p.WaitForExit()
 if($p.ExitCode -ne 0){throw $spec};return ,$m.ToArray()
}
function Extract([byte[]]$b) {
 $ms=[IO.MemoryStream]::new($b,$false);$z=[IO.Compression.ZipArchive]::new($ms,'Read');$result=''
 try {foreach($e in $z.Entries){
  if($e.FullName -notmatch '^customXml/[^/]+\.xml$'){continue}
  $r=[IO.StreamReader]::new($e.Open());try{$x=[xml]$r.ReadToEnd()}finally{$r.Dispose()}
  if($x.DocumentElement.LocalName -ne 'DataMashup'){continue}
  $d=[Convert]::FromBase64String($x.DocumentElement.InnerText);$im=[IO.MemoryStream]::new($d,8,[BitConverter]::ToUInt32($d,4));$iz=[IO.Compression.ZipArchive]::new($im,'Read')
  try {foreach($q in $iz.Entries){if($q.FullName -match '\.m$'){$r=[IO.StreamReader]::new($q.Open());try{$result += $r.ReadToEnd()}finally{$r.Dispose()}}}}finally{$iz.Dispose();$im.Dispose()}
 }}finally{$z.Dispose();$ms.Dispose()};return $result
}
if($FunctionsOnly){return}
$rows=[Collections.Generic.List[object]]::new()
function Record($label,$path,[byte[]]$b,$stamp) {
 try {
 $t=if($path -match '\.m$'){[Text.Encoding]::UTF8.GetString($b)}else{Extract $b}
 $q=[regex]::Match($t,'(?s)shared #"Availabilities Appended"\s*=.*?(?=\r?\nshared |\z)').Value
 $hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes(($t -replace '(?s)^.*?(?=section Section1;)','' -replace '\r\n',"`n").Trim())))
 $id=$rows.Count
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "$id.m"),$t)
 $rows.Add([pscustomobject]@{id=$id;version=$label;path=$path;date=$stamp;hash=$hash;filter108=($q -match '\[Resource\]\s*=\s*108');all108=@($t -split '\r?\n' | Where-Object {$_ -match '\b108\b'});queries=@([regex]::Matches($t,'(?m)^shared (.*?) =')|ForEach-Object {$_.Groups[1].Value});availabilityQuery=$q})
 }catch{$rows.Add([pscustomobject]@{version=$label;path=$path;error=$_.Exception.Message})}
}
$paths=@(rg --files -uuu CLIENT | Where-Object {$_ -match 'Effort-All.*(xlsx|\.m)' -and $_ -notmatch '~\$|\.hyper$'})
$paths+=@(rg --files -uuu CLIENT/DATExx-Whiddon/UNITS/Unit1 | Where-Object {$_ -match '[\\/]Effort\.xlsx' -and $_ -notmatch '~\$'})
foreach($path in $paths){Record 'disk' $path (ReadBytes (Join-Path $root $path)) (Get-Item -LiteralPath $path).LastWriteTime.ToString('o')}
$commits=@(git log --all --format=%H -- '*Effort-All*' 'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Effort.xlsx*')
$seen=@{}
foreach($c in $commits){
 $date=git show -s --format=%aI $c
 $entries=git ls-tree -r $c -- CLIENT
 foreach($e in $entries){
 if($e -notmatch '^\d+ blob ([a-f0-9]+)\t(.+)$'){continue};$blob=$Matches[1];$path=$Matches[2]
 if($path -notmatch 'Effort-All.*(xlsx|\.m)' -and $path -notmatch 'DATExx-Whiddon/UNITS/Unit1/2\. Calculations/Effort\.xlsx'){continue}
 if($path -match '~\$|\.hyper$'){continue};$key="$blob|$path";if($seen.ContainsKey($key)){continue};$seen[$key]=$true
 Record $c.Substring(0,7) $path (GitBytes $blob) $date
 }
}
$rows | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $PSScriptRoot 'audit.json')
$rows | Select-Object id,version,date,filter108,path,error | Format-Table -AutoSize | Out-String -Width 240
