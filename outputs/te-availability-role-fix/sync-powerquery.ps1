[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Audit', 'Sync')]
    [string] $Mode,

    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string] $SourcePath,

    [Parameter(Mandatory = $true)]
    [string] $EvidencePath,

    [string] $BackupPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-Sha256([string] $Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Normalize-Formula([string] $Value) {
    if ($null -eq $Value) { return '' }
    (($Value -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Read-MDefinitions([string] $Path) {
    $text = [IO.File]::ReadAllText($Path)
    if ($text -notmatch '(?m)^section\s+Section1\s*;') {
        throw 'The approved source has no section Section1 marker.'
    }

    $pattern = '(?m)^shared[ \t]+(?:(?:#"(?<quoted>(?:""|[^"])*)")|(?<plain>[A-Za-z_][A-Za-z0-9_]*))[ \t]*='
    $matches = [regex]::Matches($text, $pattern)
    if ($matches.Count -eq 0) { throw 'No shared Power Query definitions were found.' }

    $definitions = [ordered]@{}
    foreach ($match in $matches) {
        $name = if ($match.Groups['quoted'].Success) {
            $match.Groups['quoted'].Value -replace '""', '"'
        } else {
            $match.Groups['plain'].Value
        }
        if ($definitions.Contains($name)) { throw "Duplicate shared query name: $name" }

        $start = $match.Index + $match.Length
        $index = $start
        $inString = $false
        $inLineComment = $false
        $blockCommentDepth = 0
        $end = -1
        while ($index -lt $text.Length) {
            $current = $text[$index]
            $next = if ($index + 1 -lt $text.Length) { $text[$index + 1] } else { [char]0 }

            if ($inLineComment) {
                if ($current -eq "`n" -or $current -eq "`r") { $inLineComment = $false }
                $index++
                continue
            }
            if ($blockCommentDepth -gt 0) {
                if ($current -eq '/' -and $next -eq '*') { $blockCommentDepth++; $index += 2; continue }
                if ($current -eq '*' -and $next -eq '/') { $blockCommentDepth--; $index += 2; continue }
                $index++
                continue
            }
            if ($inString) {
                if ($current -eq '"' -and $next -eq '"') { $index += 2; continue }
                if ($current -eq '"') { $inString = $false }
                $index++
                continue
            }

            if ($current -eq '/' -and $next -eq '/') { $inLineComment = $true; $index += 2; continue }
            if ($current -eq '/' -and $next -eq '*') { $blockCommentDepth = 1; $index += 2; continue }
            if ($current -eq '"') { $inString = $true; $index++; continue }
            if ($current -eq ';') { $end = $index; break }
            $index++
        }
        if ($end -lt 0) { throw "No terminating semicolon was found for shared query '$name'." }
        $formula = $text.Substring($start, $end - $start).Trim()
        if ([string]::IsNullOrWhiteSpace($formula)) { throw "Shared query '$name' has an empty formula." }
        $definitions[$name] = $formula
    }
    $definitions
}

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Read-WorkbookQueries([string] $Path, [bool] $ReadOnly) {
    $excel = $null
    $workbook = $null
    $queries = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3
        $workbook = $excel.Workbooks.Open($Path, 0, $ReadOnly)
        if (-not $ReadOnly -and $workbook.ReadOnly) { throw 'Excel opened the workbook read-only.' }
        $queries = $workbook.Queries
        $result = [ordered]@{}
        for ($index = 1; $index -le $queries.Count; $index++) {
            $query = $null
            try {
                $query = $queries.Item($index)
                $result[[string] $query.Name] = [string] $query.Formula
            }
            finally { Release-ComObject $query }
        }
        $result
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        Release-ComObject $queries
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Compare-Definitions($SourceDefinitions, $WorkbookDefinitions) {
    $sourceNames = @($SourceDefinitions.Keys)
    $workbookNames = @($WorkbookDefinitions.Keys)
    $sourceOnly = @($sourceNames | Where-Object { -not $WorkbookDefinitions.Contains($_) })
    $workbookOnly = @($workbookNames | Where-Object { -not $SourceDefinitions.Contains($_) })
    $different = @($sourceNames | Where-Object {
        $WorkbookDefinitions.Contains($_) -and
        (Normalize-Formula $SourceDefinitions[$_]) -cne (Normalize-Formula $WorkbookDefinitions[$_])
    })
    [ordered]@{
        sourceQueryCount = $sourceNames.Count
        workbookQueryCount = $workbookNames.Count
        sourceOnly = $sourceOnly
        workbookOnly = $workbookOnly
        different = $different
        matches = $sourceOnly.Count -eq 0 -and $workbookOnly.Count -eq 0 -and $different.Count -eq 0
    }
}

function Sync-WorkbookQueries([string] $Path, $SourceDefinitions) {
    $excel = $null
    $workbook = $null
    $queries = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = 3
        $workbook = $excel.Workbooks.Open($Path, 0, $false)
        if ($workbook.ReadOnly) { throw 'Excel opened the workbook read-only.' }
        $queries = $workbook.Queries

        $existing = @{}
        for ($index = 1; $index -le $queries.Count; $index++) {
            $query = $null
            try {
                $query = $queries.Item($index)
                $existing[[string] $query.Name] = $true
            }
            finally { Release-ComObject $query }
        }

        $updated = [Collections.Generic.List[string]]::new()
        $added = [Collections.Generic.List[string]]::new()
        foreach ($name in $SourceDefinitions.Keys) {
            $query = $null
            try {
                if ($existing.ContainsKey($name)) {
                    $query = $queries.Item($name)
                    if ((Normalize-Formula ([string] $query.Formula)) -cne
                        (Normalize-Formula ([string] $SourceDefinitions[$name]))) {
                        $query.Formula = [string] $SourceDefinitions[$name]
                        $updated.Add($name)
                    }
                }
                else {
                    $query = $queries.Add($name, [string] $SourceDefinitions[$name])
                    $added.Add($name)
                }
            }
            finally { Release-ComObject $query }
        }

        $workbook.Save()
        [ordered]@{ updated = @($updated); added = @($added) }
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
        if ($null -ne $excel) { $excel.Quit() }
        Release-ComObject $queries
        Release-ComObject $workbook
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSource = (Resolve-Path -LiteralPath $SourcePath).Path
$resolvedEvidence = [IO.Path]::GetFullPath($EvidencePath)
$sourceHash = Get-Sha256 $resolvedSource
$workbookHashBefore = Get-Sha256 $resolvedWorkbook
$sourceDefinitions = Read-MDefinitions $resolvedSource
$workbookDefinitionsBefore = Read-WorkbookQueries $resolvedWorkbook $true
$comparisonBefore = Compare-Definitions $sourceDefinitions $workbookDefinitionsBefore

$syncResult = $null
$resolvedBackup = $null
if ($Mode -eq 'Sync') {
    if ($comparisonBefore.workbookOnly.Count -gt 0) {
        throw 'The workbook contains queries absent from the approved source; synchronization stopped.'
    }
    if ([string]::IsNullOrWhiteSpace($BackupPath)) { throw 'BackupPath is required in Sync mode.' }
    $resolvedBackup = [IO.Path]::GetFullPath($BackupPath)
    $backupDirectory = Split-Path -Parent $resolvedBackup
    if (-not (Test-Path -LiteralPath $backupDirectory)) {
        New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
    }
    Copy-Item -LiteralPath $resolvedWorkbook -Destination $resolvedBackup -Force
    if ((Get-Sha256 $resolvedBackup) -ne $workbookHashBefore) { throw 'The recoverable workbook backup hash does not match.' }
    $syncResult = Sync-WorkbookQueries $resolvedWorkbook $sourceDefinitions
}

$workbookDefinitionsAfter = Read-WorkbookQueries $resolvedWorkbook $true
$comparisonAfter = Compare-Definitions $sourceDefinitions $workbookDefinitionsAfter
$workbookHashAfter = Get-Sha256 $resolvedWorkbook
if ($Mode -eq 'Sync' -and -not $comparisonAfter.matches) {
    throw 'Post-sync query comparison failed.'
}

$evidence = [ordered]@{
    mode = $Mode
    workbook = $resolvedWorkbook
    source = $resolvedSource
    sourceSHA256 = $sourceHash
    workbookSHA256Before = $workbookHashBefore
    workbookSHA256After = $workbookHashAfter
    workbookModifiedAfter = (Get-Item -LiteralPath $resolvedWorkbook).LastWriteTime.ToString('o')
    backup = $resolvedBackup
    comparisonBefore = $comparisonBefore
    syncResult = $syncResult
    comparisonAfter = $comparisonAfter
    recordedAt = (Get-Date).ToString('o')
}
$evidenceDirectory = Split-Path -Parent $resolvedEvidence
if (-not (Test-Path -LiteralPath $evidenceDirectory)) {
    New-Item -ItemType Directory -Path $evidenceDirectory -Force | Out-Null
}
$evidence | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resolvedEvidence -Encoding utf8
$evidence | ConvertTo-Json -Depth 8
