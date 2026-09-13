param([string] $SnapshotName = 'baseline')
$ErrorActionPreference = 'Stop'
if ($SnapshotName -notmatch '^[a-z0-9-]+$') { throw 'Use a simple snapshot label.' }
# Read only the exact approved saved workbook. Sharing access never saves or refreshes Excel.
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$workbookRelative = 'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx'
$sourceRelative = $workbookRelative + '_PowerQuery.m'
$workbookPath = Join-Path $repoRoot $workbookRelative
$sourcePath = Join-Path $repoRoot $sourceRelative
function Read-SavedBytes([string] $path) {
    $stream = [IO.FileStream]::new($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $memory = [IO.MemoryStream]::new()
    try { $stream.CopyTo($memory); return ,$memory.ToArray() }
    finally { $stream.Dispose(); $memory.Dispose() }
}
function Get-BytesHash([byte[]] $bytes) {
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}
$workbookBytes = Read-SavedBytes $workbookPath
$sourceBytes = [IO.File]::ReadAllBytes($sourcePath)
$memory = [IO.MemoryStream]::new($workbookBytes, $false)
$archive = [IO.Compression.ZipArchive]::new($memory, [IO.Compression.ZipArchiveMode]::Read)
$sections = @()
try {
    foreach ($entry in $archive.Entries) {
        if ($entry.FullName -notmatch '^customXml/[^/]+\.xml$') { continue }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        if ($xml.DocumentElement.LocalName -ne 'DataMashup') { continue }
        $bytes = [Convert]::FromBase64String($xml.DocumentElement.InnerText)
        $length = [BitConverter]::ToUInt32($bytes, 4)
        $packageMemory = [IO.MemoryStream]::new($bytes, 8, $length)
        $package = [IO.Compression.ZipArchive]::new($packageMemory, [IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($member in $package.Entries) {
                if ($member.FullName -notmatch '\.m$') { continue }
                $mReader = [IO.StreamReader]::new($member.Open())
                try { $sections += [pscustomobject]@{ Member = $member.FullName; Text = $mReader.ReadToEnd() } }
                finally { $mReader.Dispose() }
            }
        } finally { $package.Dispose(); $packageMemory.Dispose() }
    }
} finally { $archive.Dispose(); $memory.Dispose() }
if ($sections.Count -ne 1) { throw "Expected one embedded M section, found $($sections.Count)." }
$bookHash = Get-BytesHash $workbookBytes
$sourceHash = Get-BytesHash $sourceBytes
if ($bookHash -ne (Get-BytesHash (Read-SavedBytes $workbookPath))) { throw 'Saved workbook changed during extraction.' }
if ($sourceHash -ne (Get-BytesHash ([IO.File]::ReadAllBytes($sourcePath)))) { throw 'M source changed during extraction.' }
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "AIN-B.embedded.$SnapshotName.m"), $sections[0].Text, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllBytes((Join-Path $PSScriptRoot "AIN-B.source.$SnapshotName.m"), $sourceBytes)
$evidence = [ordered]@{
    workbook = $workbookRelative
    source = $sourceRelative
    workbookSHA256 = $bookHash
    sourceSHA256 = $sourceHash
    embeddedMember = $sections[0].Member
    extraction = 'Read-only .NET ZIP/XML extraction from stable saved bytes; workbook never written'
    workbookUnchanged = $true
    sourceUnchanged = $true
    unsavedExcelChangesIncluded = $false
    extractedAt = (Get-Date).ToString('o')
}
$evidenceFile = if ($SnapshotName -eq 'baseline') { 'source-gate.json' } else { "source-gate.$SnapshotName.json" }
$evidence | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot $evidenceFile) -Encoding utf8
$evidence | ConvertTo-Json
