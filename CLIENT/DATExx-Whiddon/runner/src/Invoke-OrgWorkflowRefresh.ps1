param(
    [Parameter(Mandatory = $true)]
    [string] $RunRoot,

    [Parameter(Mandatory = $true)]
    [string] $SequencePath,

    [int] $StartAtSequence = 0,

    [int] $EndAtSequence = 0,

    [string] $StartAtWorkbook = $null,

    [string] $VisibleOverride = "false",

    [string] $TimeoutMinutesOverride = "30",

    [switch] $ValidateSelectionOnly,

    [string] $LogPath = $null,

    [string] $StatusPath = $null,

    [string] $StopRequestPath = $null
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'RefreshRunGate.ps1')
$script:BatchGateLease = $null

$script:WorkflowName = "ResidentialCare Date Refresh"
$script:LogFile = $LogPath
$script:StatusFile = $StatusPath
$script:ExitCode = 1

function Get-NormalizedFullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    return [System.IO.Path]::GetFullPath($Path).TrimEnd([char[]] @("\", "/"))
}

function Test-IsUnderRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Root
    )

    $normalizedPath = Get-NormalizedFullPath -Path $Path
    $normalizedRoot = Get-NormalizedFullPath -Path $Root

    return (
        $normalizedPath.Equals($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $normalizedPath.StartsWith($normalizedRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
    )
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string] $Level = "INFO"
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line

    if (-not [string]::IsNullOrWhiteSpace($script:LogFile)) {
        Add-Content -LiteralPath $script:LogFile -Value $line
    }
}

function Write-CurrentStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string] $State,

        [string] $CurrentWorkbook = $null,

        [string] $Message = $null
    )

    if ([string]::IsNullOrWhiteSpace($script:StatusFile)) {
        return
    }

    $lines = @(
        "state=$State",
        "timestamp=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "workflow=$script:WorkflowName",
        "logFile=$script:LogFile",
        "stage=Org calculation"
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentWorkbook)) {
        $lines += "currentWorkbook=$CurrentWorkbook"
    }

    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $lines += "message=$Message"
    }

    Set-Content -LiteralPath $script:StatusFile -Value $lines
}

function Test-ExcelTempLockExists {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $folder = Split-Path -Path $Path -Parent
    $fileName = Split-Path -Path $Path -Leaf
    $lockPath = Join-Path $folder ("~$" + $fileName)
    return (Test-Path -LiteralPath $lockPath -PathType Leaf)
}

function Test-WorkbookWritable {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
        $stream.Close()
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) {
            $stream.Close()
            $stream.Dispose()
        }
    }
}

function Assert-RequiredFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Root,

        [Parameter(Mandatory = $true)]
        [string] $RelativePath
    )

    $path = Join-Path $Root $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Container)) {
        $script:ExitCode = 10
        throw "Required folder is missing: $RelativePath"
    }

    return (Resolve-Path -LiteralPath $path -ErrorAction Stop).Path
}

function Assert-SafeWorkbookRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RelativePath,

        [Parameter(Mandatory = $true)]
        [string] $Label
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        $script:ExitCode = 11
        throw "$Label is blank."
    }

    if ([System.IO.Path]::IsPathRooted($RelativePath) -or $RelativePath -match '(^|[\\/])\.\.([\\/]|$)') {
        $script:ExitCode = 11
        throw "$Label must be a safe Date-root-relative path: $RelativePath"
    }

    $fileName = [System.IO.Path]::GetFileName($RelativePath)
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $script:ExitCode = 11
        throw "$Label must include a workbook file name: $RelativePath"
    }

    if ([System.IO.Path]::GetExtension($fileName) -ine ".xlsx") {
        $script:ExitCode = 11
        throw "$Label must be an .xlsx workbook: $RelativePath"
    }

    if ($fileName -like '~$*' -or $fileName -match '\.backup\.') {
        $script:ExitCode = 11
        throw "$Label is an excluded temp or backup workbook: $RelativePath"
    }
}

function Resolve-WorkbookPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $ResolvedRunRoot,

        [Parameter(Mandatory = $true)]
        [string] $RelativePath,

        [Parameter(Mandatory = $true)]
        [string] $Label
    )

    Assert-SafeWorkbookRelativePath -RelativePath $RelativePath -Label $Label

    $candidatePath = Join-Path $ResolvedRunRoot $RelativePath
    if (-not (Test-IsUnderRoot -Path $candidatePath -Root $ResolvedRunRoot)) {
        $script:ExitCode = 11
        throw "$Label escapes Date root: $RelativePath"
    }

    if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
        $script:ExitCode = 12
        throw "$Label is missing: $RelativePath"
    }

    $resolvedPath = (Resolve-Path -LiteralPath $candidatePath -ErrorAction Stop).Path
    if (Test-ExcelTempLockExists -Path $resolvedPath) {
        $script:ExitCode = 12
        throw "$Label appears open in Excel because a temporary lock file exists: $RelativePath"
    }

    if (-not (Test-WorkbookWritable -Path $resolvedPath)) {
        $script:ExitCode = 12
        throw "$Label is not writable for refresh: $RelativePath"
    }

    return $resolvedPath
}

function Resolve-OrgWorkbookEntries {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $SequenceData,

        [Parameter(Mandatory = $true)]
        [string] $ResolvedRunRoot
    )

    $approvedPaths = @(
        "2. Calculations/E-O-I/StafMasterList-All.xlsx",
        "2. Calculations/E-O-I/Effort-All.xlsx",
        "2. Calculations/E-O-I/EffortOutcomes.xlsx",
        "2. Calculations/E-O-I/Inefficiencies.xlsx",
        "2. Calculations/Cost/Cost..xlsx",
        "2. Calculations/EOW/EffortOutcomeLogXY.xlsx",
        "2. Calculations/EOW/2DRead.xlsx",
        "2. Calculations/Tableau Connection.xlsx"
    )

    if (-not $SequenceData.ContainsKey("Workbooks")) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: missing Workbooks."
    }

    if ($SequenceData.ContainsKey("Workflow") -and -not [string]::IsNullOrWhiteSpace([string] $SequenceData["Workflow"])) {
        $script:WorkflowName = [string] $SequenceData["Workflow"]
    }

    $entries = @($SequenceData["Workbooks"])
    if ($entries.Count -ne $approvedPaths.Count) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: expected $($approvedPaths.Count) org workbooks, found $($entries.Count)."
    }

    $seenSequences = @{}
    $resolvedEntries = @()

    foreach ($entry in $entries) {
        if (-not ($entry -is [hashtable])) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: each workbook entry must be a hashtable."
        }

        foreach ($key in @("Sequence", "Path")) {
            if (-not $entry.ContainsKey($key)) {
                $script:ExitCode = 11
                throw "Sequence file is invalid: workbook entry is missing $key."
            }
        }

        $sequence = [int] $entry["Sequence"]
        $relativePath = [string] $entry["Path"]

        if ($sequence -lt 1 -or $sequence -gt $approvedPaths.Count) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: sequence $sequence is outside the approved org range."
        }

        if ($seenSequences.ContainsKey($sequence)) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: duplicate sequence number $sequence."
        }
        $seenSequences[$sequence] = $true

        $expectedPath = $approvedPaths[$sequence - 1]
        if ($relativePath -cne $expectedPath) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: sequence $sequence must be '$expectedPath', not '$relativePath'."
        }

        $resolvedPath = Resolve-WorkbookPath -ResolvedRunRoot $ResolvedRunRoot -RelativePath $relativePath -Label "Org workbook sequence $sequence"
        $resolvedEntries += [pscustomobject] @{
            Sequence = $sequence
            RelativePath = $relativePath
            FileName = [System.IO.Path]::GetFileName($relativePath)
            QualifiedName = $relativePath
            Path = $resolvedPath
        }
    }

    return @($resolvedEntries | Sort-Object Sequence)
}

function ConvertTo-WorkbookMatchText {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Value
    )

    return (($Value.Trim() -replace '\\', '/') -replace '\s*/\s*', '/').ToLowerInvariant()
}

function Get-WorkbookMatchTokens {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Entry
    )

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Entry.FileName)
    return @(
        $Entry.RelativePath,
        $Entry.FileName,
        $stem
    )
}

function Join-WorkbookDisplayNames {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Entries
    )

    return (@($Entries | ForEach-Object { $_.RelativePath }) -join ', ')
}

function Resolve-OrgWorkbookStart {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Entries,

        [Parameter(Mandatory = $true)]
        [string] $WorkbookInput
    )

    $trimmedInput = $WorkbookInput.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedInput)) {
        $script:ExitCode = 11
        throw "Start workbook is blank."
    }

    $normalizedInput = ConvertTo-WorkbookMatchText -Value $trimmedInput
    $exactMatches = @($Entries | Where-Object {
        $entry = $_
        @(Get-WorkbookMatchTokens -Entry $entry | ForEach-Object { ConvertTo-WorkbookMatchText -Value ([string] $_) }) -contains $normalizedInput
    })

    if ($exactMatches.Count -eq 1) {
        return $exactMatches[0]
    }

    if ($exactMatches.Count -gt 1) {
        $script:ExitCode = 11
        throw "Start workbook '$WorkbookInput' is ambiguous. Matches: $(Join-WorkbookDisplayNames -Entries $exactMatches)."
    }

    $partialMatches = @($Entries | Where-Object {
        $entry = $_
        $match = $false
        foreach ($token in @(Get-WorkbookMatchTokens -Entry $entry)) {
            if ((ConvertTo-WorkbookMatchText -Value ([string] $token)).IndexOf($normalizedInput, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $match = $true
                break
            }
        }
        $match
    })

    if ($partialMatches.Count -eq 1) {
        return $partialMatches[0]
    }

    if ($partialMatches.Count -gt 1) {
        $script:ExitCode = 11
        throw "Start workbook '$WorkbookInput' is ambiguous. Matches: $(Join-WorkbookDisplayNames -Entries $partialMatches)."
    }

    $script:ExitCode = 11
    throw "Start workbook '$WorkbookInput' was not found in the org workbook sequence."
}

function Select-OrgWorkbookEntries {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Entries
    )

    if ($StartAtSequence -gt 0 -and -not [string]::IsNullOrWhiteSpace($StartAtWorkbook)) {
        $script:ExitCode = 11
        throw "Use StartAtSequence or StartAtWorkbook, not both."
    }

    if ($EndAtSequence -gt 0 -and -not [string]::IsNullOrWhiteSpace($StartAtWorkbook)) {
        $script:ExitCode = 11
        throw "Use EndAtSequence only with StartAtSequence, not StartAtWorkbook."
    }

    if ($EndAtSequence -gt 0 -and $StartAtSequence -lt 1) {
        $script:ExitCode = 11
        throw "EndAtSequence requires StartAtSequence."
    }

    if ($StartAtSequence -gt 0) {
        if ($EndAtSequence -gt 0 -and $EndAtSequence -lt $StartAtSequence) {
            $script:ExitCode = 11
            throw "Invalid org sequence range: end sequence $EndAtSequence is before start sequence $StartAtSequence."
        }

        $selected = @($Entries | Where-Object {
            $_.Sequence -ge $StartAtSequence -and ($EndAtSequence -eq 0 -or $_.Sequence -le $EndAtSequence)
        })
        if ($selected.Count -eq 0) {
            $script:ExitCode = 11
            if ($EndAtSequence -gt 0) {
                throw "Org sequence range $StartAtSequence-$EndAtSequence does not match any configured workbook."
            }

            throw "StartAtSequence $StartAtSequence is after the end of the org workbook sequence."
        }

        if ($selected[0].Sequence -ne $StartAtSequence) {
            $script:ExitCode = 11
            throw "Start sequence $StartAtSequence was not found in the org workbook sequence."
        }

        if ($EndAtSequence -gt 0 -and $selected[-1].Sequence -ne $EndAtSequence) {
            $script:ExitCode = 11
            throw "End sequence $EndAtSequence was not found in the org workbook sequence."
        }

        return $selected
    }

    if (-not [string]::IsNullOrWhiteSpace($StartAtWorkbook)) {
        $match = Resolve-OrgWorkbookStart -Entries $Entries -WorkbookInput $StartAtWorkbook
        return @($Entries | Where-Object { $_.Sequence -ge $match.Sequence })
    }

    return $Entries
}

function Resolve-SupportFiles {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $SequenceData,

        [Parameter(Mandatory = $true)]
        [string] $ResolvedRunRoot
    )

    $approvedSupportFiles = @(
        "2. Calculations/Change/AllocationChange.xlsx",
        "2. Calculations/EOW/Grid Thresholds.xlsx",
        "UNITS/Unit1/2. Calculations/Settings Data.xlsx"
    )

    if (-not $SequenceData.ContainsKey("RequiredSupportFiles")) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: missing RequiredSupportFiles."
    }

    $supportFiles = @($SequenceData["RequiredSupportFiles"])
    if ($supportFiles.Count -ne $approvedSupportFiles.Count) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: expected $($approvedSupportFiles.Count) support files, found $($supportFiles.Count)."
    }

    $resolvedSupportFiles = @()
    foreach ($supportFile in $supportFiles) {
        $relativePath = [string] $supportFile
        if ($approvedSupportFiles -notcontains $relativePath) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: support file is not approved: $relativePath"
        }

        $resolvedPath = Resolve-WorkbookPath -ResolvedRunRoot $ResolvedRunRoot -RelativePath $relativePath -Label "Required support workbook"
        $resolvedSupportFiles += [pscustomobject] @{
            RelativePath = $relativePath
            Path = $resolvedPath
        }
    }

    return $resolvedSupportFiles
}

function Get-StopRequestMode {
    if ([string]::IsNullOrWhiteSpace($StopRequestPath) -or -not (Test-Path -LiteralPath $StopRequestPath -PathType Leaf)) {
        return $null
    }

    foreach ($line in @(Get-Content -LiteralPath $StopRequestPath -ErrorAction SilentlyContinue)) {
        if ($line -match '^mode=(.+)$') {
            $mode = $matches[1].Trim()
            if ($mode -eq "Now" -or $mode -eq "AfterCurrent") {
                return $mode
            }
        }
    }

    return "AfterCurrent"
}

function Stop-WorkflowForOperator {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Message,

        [string] $CurrentWorkbook = $null
    )

    Write-Log $Message "WARN"
    Write-CurrentStatus -State "STOPPING" -CurrentWorkbook $CurrentWorkbook -Message $Message
    Write-CurrentStatus -State "STOPPED" -CurrentWorkbook $CurrentWorkbook -Message $Message
    exit 3
}

try {
    $resolvedRunRoot = (Resolve-Path -LiteralPath $RunRoot -ErrorAction Stop).Path
    if (-not $ValidateSelectionOnly) {
        $script:BatchGateLease = Enter-RefreshRunGate -DateRoot (Find-RefreshDateRoot $resolvedRunRoot)
    }
    $resolvedSequencePath = (Resolve-Path -LiteralPath $SequencePath -ErrorAction Stop).Path

    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "2. Calculations" | Out-Null
    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "2. Calculations/E-O-I" | Out-Null
    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "2. Calculations/EOW" | Out-Null
    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "2. Calculations/Cost" | Out-Null
    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "2. Calculations/Change" | Out-Null
    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "UNITS" | Out-Null
    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "runner" | Out-Null
    Assert-RequiredFolder -Root $resolvedRunRoot -RelativePath "runner/src" | Out-Null

    if (-not (Test-IsUnderRoot -Path $resolvedSequencePath -Root $resolvedRunRoot)) {
        $script:ExitCode = 11
        throw "Sequence file must be under the Date root."
    }

    $excelRefreshScript = Join-Path $PSScriptRoot "Invoke-OrgExcelWorkbookRefresh.ps1"
    if (-not (Test-Path -LiteralPath $excelRefreshScript -PathType Leaf)) {
        $script:ExitCode = 10
        throw "Required refresh engine is missing: runner/src/Invoke-OrgExcelWorkbookRefresh.ps1"
    }

    $sequenceData = Import-PowerShellDataFile -LiteralPath $resolvedSequencePath
    if (-not ($sequenceData -is [hashtable])) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: expected hashtable root."
    }

    $entries = Resolve-OrgWorkbookEntries -SequenceData $sequenceData -ResolvedRunRoot $resolvedRunRoot
    $selectedEntries = Select-OrgWorkbookEntries -Entries $entries
    $supportFiles = Resolve-SupportFiles -SequenceData $sequenceData -ResolvedRunRoot $resolvedRunRoot

    if ($ValidateSelectionOnly) {
        Write-Host ("Org workbook sequence OK: {0} selected workbook(s) of {1} total." -f $selectedEntries.Count, $entries.Count)
        foreach ($entry in $selectedEntries) {
            Write-Host ("[{0}] {1}" -f $entry.Sequence, $entry.RelativePath)
        }

        Write-Host ""
        Write-Host ("Required support files OK: {0} file(s)." -f $supportFiles.Count)
        foreach ($supportFile in $supportFiles) {
            Write-Host ("[support] {0}" -f $supportFile.RelativePath)
        }

        Write-Host ""
        Write-Host "Date-level org stage ready."
        exit 0
    }

    Write-CurrentStatus -State "RUNNING" -Message "Starting org calculation stage."
    Write-Log "Starting ResidentialCare Date Refresh org calculation stage."
    Write-Log "Date root: $resolvedRunRoot"
    Write-Log "Sequence file: $resolvedSequencePath"
    Write-Log ("Validated {0} org workbook(s); {1} selected for this run." -f $entries.Count, $selectedEntries.Count)
    foreach ($entry in $selectedEntries) {
        Write-Log ("[{0}] {1}" -f $entry.Sequence, $entry.RelativePath)
    }

    foreach ($supportFile in $supportFiles) {
        Write-Log ("Support file validated: {0}" -f $supportFile.RelativePath)
    }

    foreach ($entry in $selectedEntries) {
        $stopModeBefore = Get-StopRequestMode
        if (-not [string]::IsNullOrWhiteSpace($stopModeBefore)) {
            Stop-WorkflowForOperator -Message "Operator stop request observed before starting $($entry.QualifiedName)." -CurrentWorkbook $entry.QualifiedName
        }

        Write-CurrentStatus -State "RUNNING" -CurrentWorkbook $entry.QualifiedName
        Write-Log ("Refreshing org sequence {0}: {1}" -f $entry.Sequence, $entry.QualifiedName)

        $arguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $excelRefreshScript,
            "-WorkbookPath", $entry.Path,
            "-WorkbookName", $entry.QualifiedName,
            "-LogPath", $script:LogFile,
            "-StatusPath", $script:StatusFile,
            "-StatusPrefix", ("[Org {0}/{1}] {2}:" -f $entry.Sequence, $entries.Count, $entry.FileName),
            "-StopRequestPath", $StopRequestPath,
            "-VisibleOverride", $VisibleOverride,
            "-TimeoutMinutesOverride", $TimeoutMinutesOverride,
            "-SkipAsyncWait", "true",
            "-CleanupGhostExcelProcessesOverride", "false",
            "-ForceCloseExcelProcessesOverride", "false"
        )

        & pwsh @arguments
        $childExitCode = $LASTEXITCODE

        if ($childExitCode -eq 3) {
            Stop-WorkflowForOperator -Message "Operator stop request observed while refreshing $($entry.QualifiedName)." -CurrentWorkbook $entry.QualifiedName
        }

        if ($childExitCode -ne 0) {
            $script:ExitCode = $childExitCode
            throw "Org workbook refresh failed for $($entry.QualifiedName). Exit code: $childExitCode"
        }

        Write-Log ("Org workbook refresh succeeded: {0}" -f $entry.QualifiedName)

        $stopModeAfter = Get-StopRequestMode
        if (-not [string]::IsNullOrWhiteSpace($stopModeAfter)) {
            Stop-WorkflowForOperator -Message "Operator stop request observed after completing $($entry.QualifiedName)." -CurrentWorkbook $entry.QualifiedName
        }
    }

    Write-Log "ResidentialCare Date Refresh org calculation stage succeeded."
    Write-CurrentStatus -State "SUCCEEDED" -Message "Org calculation stage completed."
    exit 0
}
catch {
    $message = $_.Exception.Message
    if ($script:ExitCode -eq 0) {
        $script:ExitCode = 1
    }

    if (-not $ValidateSelectionOnly) {
        Write-Log $message "ERROR"
        Write-CurrentStatus -State "FAILED" -Message $message
    }
    else {
        Write-Error $message
    }

    exit $script:ExitCode
}
finally {
    if ($null -ne $script:BatchGateLease) { $script:BatchGateLease.Dispose() }
}
