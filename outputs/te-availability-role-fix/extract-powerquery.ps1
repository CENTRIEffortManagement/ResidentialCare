[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputMPath,

    [Parameter(Mandatory = $true)]
    [string] $EvidencePath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Read-StableBytes([string] $Path) {
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $memory = [IO.MemoryStream]::new()
    try {
        $stream.CopyTo($memory)
        return ,$memory.ToArray()
    }
    finally {
        $stream.Dispose()
        $memory.Dispose()
    }
}

function Get-BytesHash([byte[]] $Bytes) {
    [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($Bytes))
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedOutput = [IO.Path]::GetFullPath($OutputMPath)
$resolvedEvidence = [IO.Path]::GetFullPath($EvidencePath)
$workbookBytes = Read-StableBytes $resolvedWorkbook
$workbookHash = Get-BytesHash $workbookBytes
$outerMemory = [IO.MemoryStream]::new($workbookBytes, $false)
$outerArchive = [IO.Compression.ZipArchive]::new($outerMemory, [IO.Compression.ZipArchiveMode]::Read)
$sections = [Collections.Generic.List[object]]::new()
try {
    foreach ($entry in $outerArchive.Entries) {
        if ($entry.FullName -notmatch '^customXml/[^/]+\.xml$') { continue }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $xml = [xml] $reader.ReadToEnd() }
        finally { $reader.Dispose() }
        if ($xml.DocumentElement.LocalName -ne 'DataMashup') { continue }

        $mashupBytes = [Convert]::FromBase64String($xml.DocumentElement.InnerText)
        $packageLength = [BitConverter]::ToUInt32($mashupBytes, 4)
        $packageMemory = [IO.MemoryStream]::new($mashupBytes, 8, $packageLength)
        $packageArchive = [IO.Compression.ZipArchive]::new($packageMemory, [IO.Compression.ZipArchiveMode]::Read)
        try {
            foreach ($member in $packageArchive.Entries) {
                if ($member.FullName -notmatch '\.m$') { continue }
                $mReader = [IO.StreamReader]::new($member.Open())
                try {
                    $sections.Add([pscustomobject]@{
                        OuterMember = $entry.FullName
                        InnerMember = $member.FullName
                        Text = $mReader.ReadToEnd()
                    })
                }
                finally { $mReader.Dispose() }
            }
        }
        finally {
            $packageArchive.Dispose()
            $packageMemory.Dispose()
        }
    }
}
finally {
    $outerArchive.Dispose()
    $outerMemory.Dispose()
}

if ($sections.Count -ne 1) { throw "Expected one embedded M section, found $($sections.Count)." }
if ($workbookHash -ne (Get-BytesHash (Read-StableBytes $resolvedWorkbook))) {
    throw 'Workbook changed during read-only extraction.'
}

$outputDirectory = Split-Path -Parent $resolvedOutput
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}
[IO.File]::WriteAllText($resolvedOutput, $sections[0].Text, [Text.UTF8Encoding]::new($false))
$evidence = [ordered]@{
    workbook = $resolvedWorkbook
    workbookSHA256 = $workbookHash
    workbookModified = (Get-Item -LiteralPath $resolvedWorkbook).LastWriteTime.ToString('o')
    outerMember = $sections[0].OuterMember
    innerMember = $sections[0].InnerMember
    extractedM = $resolvedOutput
    extractedMSHA256 = (Get-FileHash -LiteralPath $resolvedOutput -Algorithm SHA256).Hash
    workbookUnchanged = $true
    extractedAt = (Get-Date).ToString('o')
}
$evidence | ConvertTo-Json | Set-Content -LiteralPath $resolvedEvidence -Encoding utf8
$evidence | ConvertTo-Json
