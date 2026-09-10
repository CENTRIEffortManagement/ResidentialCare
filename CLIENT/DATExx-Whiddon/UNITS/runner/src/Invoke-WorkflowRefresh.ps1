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

    [switch] $ValidateSelectionOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:WorkflowName = "ResidentialCare All Units Refresh"
$script:LogFile = $null
$script:StatusFile = $null
$script:ExitCode = 1
$script:RunLockStream = $null
$script:RunLockPath = $null

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
        "logFile=$script:LogFile"
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentWorkbook)) {
        $lines += "currentWorkbook=$CurrentWorkbook"
    }

    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $lines += "message=$Message"
    }

    Set-Content -LiteralPath $script:StatusFile -Value $lines
}

function Acquire-RunLock {
    param(
        [Parameter(Mandatory = $true)]
        [string] $LockPath
    )

    $lockFolder = Split-Path -Path $LockPath -Parent
    if (-not (Test-Path -LiteralPath $lockFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $lockFolder -Force | Out-Null
    }

    try {
        $stream = [System.IO.File]::Open(
            $LockPath,
            [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $stream.SetLength(0)

        $writer = [System.IO.StreamWriter]::new($stream, [System.Text.Encoding]::UTF8, 1024, $true)
        $writer.WriteLine("workflow=$script:WorkflowName")
        $writer.WriteLine("pid=$PID")
        $writer.WriteLine("startedAt=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $writer.Flush()
        $writer.Dispose()

        $script:RunLockStream = $stream
        $script:RunLockPath = $LockPath
    }
    catch [System.IO.IOException] {
        $script:ExitCode = 13
        throw "Another ResidentialCare All Units Refresh appears to be running. Lock file is in use: $LockPath"
    }
}

function Release-RunLock {
    if ($null -ne $script:RunLockStream) {
        $script:RunLockStream.Dispose()
        $script:RunLockStream = $null
    }

    if (-not [string]::IsNullOrWhiteSpace($script:RunLockPath) -and (Test-Path -LiteralPath $script:RunLockPath -PathType Leaf)) {
        Remove-Item -LiteralPath $script:RunLockPath -Force -ErrorAction SilentlyContinue
    }

    $script:RunLockPath = $null
}

function Get-DiscoveredUnits {
    param(
        [Parameter(Mandatory = $true)]
        [string] $ResolvedRunRoot
    )

    $units = @()
    foreach ($item in @(Get-ChildItem -LiteralPath $ResolvedRunRoot -Directory -Force -ErrorAction Stop)) {
        if ($item.Name -match '^Unit(\d+)$') {
            $units += [pscustomobject] @{
                Name = $item.Name
                Number = [int] $matches[1]
                Path = (Resolve-Path -LiteralPath $item.FullName -ErrorAction Stop).Path
            }
        }
    }

    $sortedUnits = @($units | Sort-Object Number, Name)
    if ($sortedUnits.Count -eq 0) {
        $script:ExitCode = 10
        throw "No unit folders matching ^Unit\d+$ were found under $ResolvedRunRoot."
    }

    return $sortedUnits
}

function Assert-RequiredFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Root,

        [Parameter(Mandatory = $true)]
        [string] $RelativePath,

        [Parameter(Mandatory = $true)]
        [string] $UnitName
    )

    $path = Join-Path $Root $RelativePath
    if (-not (Test-IsUnderRoot -Path $path -Root $Root)) {
        $script:ExitCode = 10
        throw "Required folder path escapes $UnitName root: $RelativePath"
    }

    if (-not (Test-Path -LiteralPath $path -PathType Container)) {
        $script:ExitCode = 10
        throw "Required folder is missing for $UnitName`: $RelativePath"
    }

    return (Resolve-Path -LiteralPath $path -ErrorAction Stop).Path
}

function Resolve-PerUnitWorkbookEntries {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $SequenceData
    )

    if (-not $SequenceData.ContainsKey("Workbooks")) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: missing Workbooks."
    }

    if ($SequenceData.ContainsKey("Workflow") -and -not [string]::IsNullOrWhiteSpace([string] $SequenceData["Workflow"])) {
        $script:WorkflowName = [string] $SequenceData["Workflow"]
    }

    $allowedRelativePaths = @(
        "1. Input/1-AllocationExtracted.xlsx",
        "1. Input/2-DemandExtract.xlsx",
        "2. Calculations/Settings Data.xlsx",
        "2. Calculations/Intervals.xlsx",
        "2. Calculations/DemandIntervals.xlsx",
        "2. Calculations/Demand.xlsx",
        "2. Calculations/Capacity-ShiftAvailability.xlsx",
        "2. Calculations/StaffListMaster.xlsx",
        "2. Calculations/Shifts.xlsx",
        "2. Calculations/AllocationByShiftAverage.xlsx",
        "2. Calculations/Allocation.xlsx",
        "2. Calculations/Shift-StaffDistribution.xlsx",
        "2. Calculations/MutliRoleCheck.xlsx",
        "2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx",
        "2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx",
        "2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx",
        "2. Calculations/AINC4/CapacityDistrib(A.1)-shifts.xlsx",
        "2. Calculations/AINC4/CapacityDistrib(A.2)-shifts.xlsx",
        "2. Calculations/AINC4/CapacityDistrib(B)-shifts.xlsx",
        "2. Calculations/RN/CapacityDistrib(A.1)-shifts.xlsx",
        "2. Calculations/RN/CapacityDistrib(A.2)-shifts.xlsx",
        "2. Calculations/RN/CapacityDistrib(B)-shifts.xlsx",
        "2. Calculations/Capacity.xlsx",
        "2. Calculations/Effort.xlsx"
    )

    $entries = @($SequenceData["Workbooks"])
    if ($entries.Count -ne 24) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: expected exactly 24 workbooks, found $($entries.Count)."
    }

    $seenSequences = @{}
    $seenFiles = @{}
    $resolvedEntries = @()

    foreach ($entry in $entries) {
        if (-not ($entry -is [hashtable])) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: each workbook entry must be a hashtable."
        }

        foreach ($key in @("Sequence", "Folder", "FileName")) {
            if (-not $entry.ContainsKey($key)) {
                $script:ExitCode = 11
                throw "Sequence file is invalid: workbook entry is missing $key."
            }
        }

        $sequence = [int] $entry["Sequence"]
        $folder = [string] $entry["Folder"]
        $fileName = [string] $entry["FileName"]

        if ($sequence -lt 1 -or $sequence -gt 24) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: per-unit sequence numbers must be 1 through 24."
        }

        if ($seenSequences.ContainsKey($sequence)) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: duplicate per-unit sequence number $sequence."
        }
        $seenSequences[$sequence] = $true

        if ([string]::IsNullOrWhiteSpace($folder) -or [System.IO.Path]::IsPathRooted($folder) -or $folder -match '(^|[\\/])\.\.([\\/]|$)') {
            $script:ExitCode = 11
            throw "Sequence file is invalid: folder must be a safe unit-relative path for $fileName."
        }

        if ([System.IO.Path]::GetFileName($fileName) -ne $fileName -or [System.IO.Path]::IsPathRooted($fileName)) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: FileName must be a file name only for sequence $sequence."
        }

        if ([System.IO.Path]::GetExtension($fileName) -ine ".xlsx") {
            $script:ExitCode = 11
            throw "Sequence file is invalid: $fileName is not an .xlsx workbook."
        }

        if ($fileName -like '~$*' -or $fileName -match '\.backup\.') {
            $script:ExitCode = 11
            throw "Sequence file is invalid: excluded workbook name $fileName."
        }

        $relativePath = "$folder/$fileName"
        if ($allowedRelativePaths -notcontains $relativePath) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: workbook '$relativePath' is not in the approved ResidentialCare All Units Refresh sequence."
        }

        $fileKey = $relativePath.ToLowerInvariant()
        if ($seenFiles.ContainsKey($fileKey)) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: duplicate workbook $folder/$fileName."
        }
        $seenFiles[$fileKey] = $true

        $roleName = $null
        if ($folder -match '[\\/]') {
            $roleName = @($folder -split '[\\/]')[-1]
        }

        $resolvedEntries += [pscustomobject] @{
            LocalSequence = $sequence
            Folder = $folder
            Role = $roleName
            FileName = $fileName
            RelativePath = $relativePath
        }
    }

    foreach ($expectedSequence in 1..24) {
        if (-not $seenSequences.ContainsKey($expectedSequence)) {
            $script:ExitCode = 11
            throw "Sequence file is invalid: missing per-unit sequence number $expectedSequence."
        }
    }

    return @($resolvedEntries | Sort-Object LocalSequence)
}

function Expand-AllUnitWorkbookEntries {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Units,

        [Parameter(Mandatory = $true)]
        [object[]] $PerUnitEntries
    )

    $globalSequence = 1
    $expanded = @()

    foreach ($unit in $Units) {
        Assert-RequiredFolder -Root $unit.Path -RelativePath "1. Input" -UnitName $unit.Name | Out-Null
        Assert-RequiredFolder -Root $unit.Path -RelativePath "2. Calculations" -UnitName $unit.Name | Out-Null
        Assert-RequiredFolder -Root $unit.Path -RelativePath "2. Calculations/AIN" -UnitName $unit.Name | Out-Null
        Assert-RequiredFolder -Root $unit.Path -RelativePath "2. Calculations/AINC4" -UnitName $unit.Name | Out-Null
        Assert-RequiredFolder -Root $unit.Path -RelativePath "2. Calculations/RN" -UnitName $unit.Name | Out-Null

        foreach ($entry in $PerUnitEntries) {
            $workbookPath = Join-Path (Join-Path $unit.Path $entry.Folder) $entry.FileName
            if (-not (Test-IsUnderRoot -Path $workbookPath -Root $unit.Path)) {
                $script:ExitCode = 11
                throw "Sequence file is invalid: workbook path for $($unit.Name) escapes that unit root for $($entry.RelativePath)."
            }

            $unitDisplayPath = if (-not [string]::IsNullOrWhiteSpace([string] $entry.Role)) {
                "{0} / {1} / {2}" -f $unit.Name, $entry.Role, $entry.FileName
            }
            else {
                "{0} / {1}" -f $unit.Name, $entry.RelativePath
            }

            $expanded += [pscustomobject] @{
                Sequence = $globalSequence
                UnitName = $unit.Name
                UnitNumber = $unit.Number
                UnitRoot = $unit.Path
                LocalSequence = $entry.LocalSequence
                Folder = $entry.Folder
                Role = $entry.Role
                FileName = $entry.FileName
                RelativePath = $entry.RelativePath
                UnitRelativePath = ("{0}/{1}" -f $unit.Name, $entry.RelativePath)
                DisplayPath = $unitDisplayPath
                Path = [System.IO.Path]::GetFullPath($workbookPath)
            }

            $globalSequence++
        }
    }

    return $expanded
}

function Format-WorkbookDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Entry
    )

    return ("{0:000} [{1} #{2:00}] {3}" -f $Entry.Sequence, $Entry.UnitName, $Entry.LocalSequence, $Entry.DisplayPath)
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
    $tokens = @(
        $Entry.FileName,
        $stem,
        $Entry.RelativePath,
        ("{0}/{1}" -f $Entry.Folder, $stem),
        $Entry.UnitRelativePath,
        ("{0}/{1}/{2}" -f $Entry.UnitName, $Entry.Folder, $stem),
        ("{0}/{1}" -f $Entry.UnitName, $Entry.FileName),
        ("{0}/{1}" -f $Entry.UnitName, $stem),
        ("{0} / {1}" -f $Entry.UnitName, $Entry.FileName),
        ("{0} / {1}" -f $Entry.UnitName, $stem),
        ("{0} / {1}" -f $Entry.UnitName, $Entry.RelativePath),
        $Entry.DisplayPath
    )

    if (-not [string]::IsNullOrWhiteSpace([string] $Entry.Role)) {
        $tokens += @(
            ("{0}/{1}" -f $Entry.Role, $Entry.FileName),
            ("{0}/{1}" -f $Entry.Role, $stem),
            ("{0} / {1}" -f $Entry.Role, $Entry.FileName),
            ("{0} / {1}" -f $Entry.Role, $stem),
            ("{0}/{1}/{2}" -f $Entry.UnitName, $Entry.Role, $Entry.FileName),
            ("{0}/{1}/{2}" -f $Entry.UnitName, $Entry.Role, $stem),
            ("{0} / {1} / {2}" -f $Entry.UnitName, $Entry.Role, $Entry.FileName),
            ("{0} / {1} / {2}" -f $Entry.UnitName, $Entry.Role, $stem)
        )
    }

    return $tokens
}

function Join-WorkbookDisplayNames {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Entries
    )

    return (@($Entries | ForEach-Object { Format-WorkbookDisplayName -Entry $_ }) -join '; ')
}

function Resolve-WorkbookStart {
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
        $tokens = @(Get-WorkbookMatchTokens -Entry $entry | ForEach-Object { ConvertTo-WorkbookMatchText -Value ([string] $_) })
        $tokens -contains $normalizedInput
    })

    if ($exactMatches.Count -eq 1) {
        return $exactMatches[0]
    }

    if ($exactMatches.Count -gt 1) {
        $script:ExitCode = 11
        throw "Start workbook '$WorkbookInput' is ambiguous. Matches: $(Join-WorkbookDisplayNames -Entries $exactMatches). Use a unit-qualified path such as 'Unit2 / Settings Data.xlsx', exact path such as 'Unit2/2. Calculations/Settings Data.xlsx', or use -StartAtSequence."
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
        throw "Start workbook '$WorkbookInput' is ambiguous. Matches: $(Join-WorkbookDisplayNames -Entries $partialMatches). Use a unit-qualified path such as 'Unit2 / Settings Data.xlsx', exact path such as 'Unit2/2. Calculations/Settings Data.xlsx', or use -StartAtSequence."
    }

    $script:ExitCode = 11
    throw "Start workbook '$WorkbookInput' was not found in the configured all-units workbook sequence."
}

function Select-WorkbookEntries {
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
            throw "Invalid sequence range: end sequence $EndAtSequence is before start sequence $StartAtSequence."
        }

        $selected = @($Entries | Where-Object {
            $_.Sequence -ge $StartAtSequence -and ($EndAtSequence -eq 0 -or $_.Sequence -le $EndAtSequence)
        })

        if ($selected.Count -eq 0) {
            $script:ExitCode = 11
            if ($EndAtSequence -gt 0) {
                throw "Global sequence range $StartAtSequence-$EndAtSequence does not match any configured workbook."
            }

            throw "StartAtSequence $StartAtSequence is after the end of the all-units workbook sequence."
        }

        if ($selected[0].Sequence -ne $StartAtSequence) {
            $script:ExitCode = 11
            throw "Start sequence $StartAtSequence was not found in the configured all-units workbook sequence."
        }

        if ($EndAtSequence -gt 0 -and $selected[-1].Sequence -ne $EndAtSequence) {
            $script:ExitCode = 11
            throw "End sequence $EndAtSequence was not found in the configured all-units workbook sequence."
        }

        return $selected
    }

    if (-not [string]::IsNullOrWhiteSpace($StartAtWorkbook)) {
        $match = Resolve-WorkbookStart -Entries $Entries -WorkbookInput $StartAtWorkbook
        return @($Entries | Where-Object { $_.Sequence -ge $match.Sequence })
    }

    return $Entries
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

function Assert-SelectedWorkbookFiles {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $SelectedEntries
    )

    foreach ($entry in $SelectedEntries) {
        if (-not (Test-IsUnderRoot -Path $entry.Path -Root $entry.UnitRoot)) {
            $script:ExitCode = 12
            throw "Selected workbook path escapes $($entry.UnitName) root: $($entry.UnitRelativePath)"
        }

        if (-not (Test-Path -LiteralPath $entry.Path -PathType Leaf)) {
            $script:ExitCode = 12
            throw "Selected workbook is missing: $(Format-WorkbookDisplayName -Entry $entry) ($($entry.UnitRelativePath))"
        }

        $resolvedPath = (Resolve-Path -LiteralPath $entry.Path -ErrorAction Stop).Path
        if (-not (Test-IsUnderRoot -Path $resolvedPath -Root $entry.UnitRoot)) {
            $script:ExitCode = 12
            throw "Selected workbook resolves outside $($entry.UnitName) root: $($entry.UnitRelativePath)"
        }

        if (Test-ExcelTempLockExists -Path $resolvedPath) {
            $script:ExitCode = 12
            throw "Selected workbook appears open in Excel because a temporary lock file exists: $(Format-WorkbookDisplayName -Entry $entry) ($($entry.UnitRelativePath))"
        }

        if (-not (Test-WorkbookWritable -Path $resolvedPath)) {
            $script:ExitCode = 12
            throw "Selected workbook is not writable for refresh: $(Format-WorkbookDisplayName -Entry $entry) ($($entry.UnitRelativePath))"
        }

        $entry.Path = $resolvedPath
    }
}

function Get-StopRequestMode {
    param(
        [Parameter(Mandatory = $true)]
        [string] $StopRequestPath
    )

    if (-not (Test-Path -LiteralPath $StopRequestPath -PathType Leaf)) {
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
    Release-RunLock
    exit 3
}

try {
    $resolvedRunRoot = (Resolve-Path -LiteralPath $RunRoot -ErrorAction Stop).Path
    $resolvedSequencePath = (Resolve-Path -LiteralPath $SequencePath -ErrorAction Stop).Path

    $excelRefreshScript = Join-Path $PSScriptRoot "Invoke-ExcelWorkbookRefresh.ps1"
    if (-not (Test-Path -LiteralPath $excelRefreshScript -PathType Leaf)) {
        $script:ExitCode = 10
        throw "Required refresh engine is missing: runner/src/Invoke-ExcelWorkbookRefresh.ps1"
    }

    if (-not (Test-IsUnderRoot -Path $resolvedSequencePath -Root $resolvedRunRoot)) {
        $script:ExitCode = 10
        throw "Sequence path must live under the UNITS run root: $resolvedSequencePath"
    }

    $sequenceData = Import-PowerShellDataFile -LiteralPath $resolvedSequencePath
    if (-not ($sequenceData -is [hashtable])) {
        $script:ExitCode = 11
        throw "Sequence file is invalid: expected hashtable root."
    }

    $units = @(Get-DiscoveredUnits -ResolvedRunRoot $resolvedRunRoot)
    $perUnitEntries = @(Resolve-PerUnitWorkbookEntries -SequenceData $sequenceData)
    $entries = @(Expand-AllUnitWorkbookEntries -Units $units -PerUnitEntries $perUnitEntries)
    $selectedEntries = @(Select-WorkbookEntries -Entries $entries)
    Assert-SelectedWorkbookFiles -SelectedEntries $selectedEntries

    if ($ValidateSelectionOnly) {
        Write-Host ("Discovered unit count: {0}" -f $units.Count)
        Write-Host ("Discovered units: {0}" -f (($units | ForEach-Object { $_.Name }) -join ', '))
        Write-Host ("Per-unit workbook count: {0}" -f $perUnitEntries.Count)
        Write-Host ("Total configured workbook count: {0}" -f $entries.Count)
        Write-Host ("Selection OK: {0} workbook(s)." -f $selectedEntries.Count)
        foreach ($entry in $selectedEntries) {
            Write-Host (Format-WorkbookDisplayName -Entry $entry)
        }
        exit 0
    }

    $runLogsPath = Join-Path $resolvedRunRoot "RunLogs"
    if (-not (Test-Path -LiteralPath $runLogsPath -PathType Container)) {
        New-Item -ItemType Directory -Path $runLogsPath -Force | Out-Null
    }

    $runLockPath = Join-Path $runLogsPath "AllUnitsRefresh.lock"
    Acquire-RunLock -LockPath $runLockPath

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $script:LogFile = Join-Path $runLogsPath ("AllUnitsRefresh-{0}.log" -f $timestamp)
    $script:StatusFile = Join-Path $runLogsPath "current-status.txt"
    $latestLogPath = Join-Path $runLogsPath "latest-log.txt"
    $stopRequestPath = Join-Path $runLogsPath "stop-request.txt"

    Set-Content -LiteralPath $latestLogPath -Value $script:LogFile

    Write-Host ("Discovered unit count: {0}" -f $units.Count)
    Write-Host ("Selected workbook count: {0}" -f $selectedEntries.Count)
    Write-Host ("Refresh log path: {0}" -f $script:LogFile)
    Write-Host ("Current status path: {0}" -f $script:StatusFile)

    if (Test-Path -LiteralPath $stopRequestPath -PathType Leaf) {
        Remove-Item -LiteralPath $stopRequestPath -Force
    }

    Write-CurrentStatus -State "RUNNING"
    Write-Log "Starting ResidentialCare All Units Refresh workflow package."
    Write-Log "UNITS root: $resolvedRunRoot"
    Write-Log "Sequence file: $resolvedSequencePath"
    Write-Log ("Discovered units: {0}" -f (($units | ForEach-Object { $_.Name }) -join ', '))
    Write-Log ("Validated {0} workbook(s); {1} workbook(s) selected for this run." -f $entries.Count, $selectedEntries.Count)
    foreach ($entry in $selectedEntries) {
        Write-Log (Format-WorkbookDisplayName -Entry $entry)
    }

    foreach ($entry in $selectedEntries) {
        $displayName = Format-WorkbookDisplayName -Entry $entry
        $stopModeBefore = Get-StopRequestMode -StopRequestPath $stopRequestPath
        if (-not [string]::IsNullOrWhiteSpace($stopModeBefore)) {
            Stop-WorkflowForOperator -Message "Operator stop request observed before starting $displayName." -CurrentWorkbook $displayName
        }

        Write-CurrentStatus -State "RUNNING" -CurrentWorkbook $displayName
        Write-Log ("Refreshing {0}" -f $displayName)

        $arguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $excelRefreshScript,
            "-WorkbookPath", $entry.Path,
            "-WorkbookName", $displayName,
            "-LogPath", $script:LogFile,
            "-StatusPath", $script:StatusFile,
            "-StatusPrefix", ("[{0}/{1} {2} #{3:00}]" -f $entry.Sequence, $entries.Count, $entry.UnitName, $entry.LocalSequence),
            "-StopRequestPath", $stopRequestPath,
            "-VisibleOverride", $VisibleOverride,
            "-TimeoutMinutesOverride", $TimeoutMinutesOverride,
            "-SkipAsyncWait", "true",
            "-CleanupGhostExcelProcessesOverride", "false",
            "-ForceCloseExcelProcessesOverride", "false"
        )

        & pwsh @arguments
        $childExitCode = $LASTEXITCODE

        if ($childExitCode -eq 3) {
            Stop-WorkflowForOperator -Message "Operator stop request observed while refreshing $displayName." -CurrentWorkbook $displayName
        }

        if ($childExitCode -ne 0) {
            $script:ExitCode = $childExitCode
            throw "Workbook refresh failed for $displayName. Exit code: $childExitCode"
        }

        Write-Log ("Workbook refresh succeeded: {0}" -f $displayName)

        $stopModeAfter = Get-StopRequestMode -StopRequestPath $stopRequestPath
        if (-not [string]::IsNullOrWhiteSpace($stopModeAfter)) {
            Stop-WorkflowForOperator -Message "Operator stop request observed after completing $displayName." -CurrentWorkbook $displayName
        }
    }

    Write-Log "ResidentialCare All Units Refresh workflow succeeded."
    Write-CurrentStatus -State "SUCCEEDED"
    Release-RunLock
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
        Release-RunLock
    }
    else {
        Write-Host "ERROR: $message"
    }

    exit $script:ExitCode
}
