param([ValidateSet('availability', 'ain-b')][string] $Target, [string] $SnapshotLabel = 'before')
$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$relativeWorkbook = if ($Target -eq 'availability') {
    'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx'
} else {
    'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx'
}
$workbookPath = Join-Path $repoRoot $relativeWorkbook
$sourcePath = $workbookPath + '_PowerQuery.m'

function Read-SavedBytes([string] $path) {
    $stream = [IO.FileStream]::new($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $memory = [IO.MemoryStream]::new()
    try { $stream.CopyTo($memory); return ,$memory.ToArray() }
    finally { $stream.Dispose(); $memory.Dispose() }
}
function Hash-Bytes([byte[]] $bytes) {
    [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

$workbookBytes = Read-SavedBytes $workbookPath
$sourceBytes = [IO.File]::ReadAllBytes($sourcePath)
$outerMemory = [IO.MemoryStream]::new($workbookBytes, $false)
$outer = [IO.Compression.ZipArchive]::new($outerMemory, [IO.Compression.ZipArchiveMode]::Read)
$sections = @()
try {
    foreach ($entry in $outer.Entries) {
        if ($entry.FullName -notmatch '^customXml/[^/]+\.xml$') { continue }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        if ($xml.DocumentElement.LocalName -ne 'DataMashup') { continue }
        $mashup = [Convert]::FromBase64String($xml.DocumentElement.InnerText)
        $packageLength = [BitConverter]::ToUInt32($mashup, 4)
        $innerMemory = [IO.MemoryStream]::new($mashup, 8, $packageLength)
        $inner = [IO.Compression.ZipArchive]::new($innerMemory, [IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($member in $inner.Entries) {
                if ($member.FullName -notmatch '\.m$') { continue }
                $mReader = [IO.StreamReader]::new($member.Open())
                try { $sections += [pscustomobject]@{ Member = $member.FullName; Text = $mReader.ReadToEnd() } }
                finally { $mReader.Dispose() }
            }
        } finally { $inner.Dispose(); $innerMemory.Dispose() }
    }
} finally { $outer.Dispose(); $outerMemory.Dispose() }

if ($sections.Count -ne 1) { throw "Expected one embedded M section, found $($sections.Count)." }
$workbookHash = Hash-Bytes $workbookBytes
$sourceHash = Hash-Bytes $sourceBytes
if ($workbookHash -ne (Hash-Bytes (Read-SavedBytes $workbookPath))) { throw 'Workbook changed during extraction.' }
if ($sourceHash -ne (Hash-Bytes ([IO.File]::ReadAllBytes($sourcePath)))) { throw 'Source changed during extraction.' }

$prefix = if ($SnapshotLabel -eq 'before') { $Target } else { "$Target.$SnapshotLabel" }
$embeddedPath = Join-Path $PSScriptRoot "$prefix.embedded.$SnapshotLabel.m"
$snapshotPath = Join-Path $PSScriptRoot "$prefix.source.$SnapshotLabel.m"
[IO.File]::WriteAllText($embeddedPath, $sections[0].Text, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllBytes($snapshotPath, $sourceBytes)
$sourceText = [Text.Encoding]::UTF8.GetString($sourceBytes)
$sectionOffset = $sourceText.IndexOf('section Section1;')
if ($sectionOffset -lt 0) { throw 'Adjacent source has no section Section1 marker.' }
$normalize = { param([string] $value) ($value -replace "`r`n", "`n").TrimEnd("`r", "`n") }
$sameSection = (& $normalize $sections[0].Text) -ceq (& $normalize $sourceText.Substring($sectionOffset))
$evidence = [ordered]@{
    workbook = $relativeWorkbook
    source = $relativeWorkbook + '_PowerQuery.m'
    workbookSHA256 = $workbookHash
    sourceSHA256 = $sourceHash
    workbookModified = (Get-Item -LiteralPath $workbookPath).LastWriteTime.ToString('o')
    sourceModified = (Get-Item -LiteralPath $sourcePath).LastWriteTime.ToString('o')
    embeddedMember = $sections[0].Member
    embeddedMatchesAdjacentSource = $sameSection
    workbookUnchangedDuringExtraction = $true
    sourceUnchangedDuringExtraction = $true
}
$evidence | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot "$prefix.source-gate.json") -Encoding utf8
$evidence | ConvertTo-Json
