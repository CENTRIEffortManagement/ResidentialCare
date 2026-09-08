param(
    [switch] $RefreshAll,

    [int] $StartAtSequence = 0,

    [int] $EndAtSequence = 0,

    [string] $StartAtWorkbook = $null,

    [string] $VisibleOverride = "false",

    [string] $TimeoutMinutesOverride = "30",

    [switch] $ValidateSelectionOnly,

    [switch] $ShowStatus,

    [ValidateSet("AfterCurrent", "Now")]
    [string] $StopMode = $null
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:WorkflowName = "ResidentialCare Date Refresh"
$script:LogFile = $null
$script:StatusFile = $null
$script:ExitCode = 1
$script:RunLockStream = $null
$script:RunLockPath = $null

$runnerRoot = $PSScriptRoot
$dateRoot = Split-Path -Path $runnerRoot -Parent
$unitsRoot = Join-Path $dateRoot "UNITS"

$unitWorkflowScript = Join-Path $runnerRoot "src\Invoke-AllUnitsWorkflowRefresh.ps1"
$unitSequencePath = Join-Path $runnerRoot "ResidentialCare-UnitWorkbookSequence.psd1"
$orgWorkflowScript = Join-Path $runnerRoot "src\Invoke-OrgWorkflowRefresh.ps1"
$orgSequencePath = Join-Path $runnerRoot "ResidentialCare-OrgWorkbookSequence.psd1"

$runLogsPath = Join-Path $dateRoot "RunLogs"
$stopRequestPath = Join-Path $runLogsPath "stop-request.txt"
$unitRunLogsPath = Join-Path $unitsRoot "RunLogs"
$unitStopRequestPath = Join-Path $unitRunLogsPath "stop-request.txt"

function Assert-PackageFile {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Description
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Write-Error "Required $Description is missing: $Path"
        exit 10
    }
}

function Assert-PackageFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Description
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        Write-Error "Required $Description is missing: $Path"
        exit 10
    }
}

function Write-DateLog {
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

function Write-DateStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string] $State,

        [string] $Stage = $null,

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

    if (-not [string]::IsNullOrWhiteSpace($Stage)) {
        $lines += "stage=$Stage"
    }

    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $lines += "message=$Message"
    }

    Set-Content -LiteralPath $script:StatusFile -Value $lines
}

function Acquire-DateRunLock {
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
        throw "Another ResidentialCare Date Refresh appears to be running. Lock file is in use: $LockPath"
    }
}

function Release-DateRunLock {
    if ($null -ne $script:RunLockStream) {
        $script:RunLockStream.Dispose()
        $script:RunLockStream = $null
    }

    if (-not [string]::IsNullOrWhiteSpace($script:RunLockPath) -and (Test-Path -LiteralPath $script:RunLockPath -PathType Leaf)) {
        Remove-Item -LiteralPath $script:RunLockPath -Force -ErrorAction SilentlyContinue
    }

    $script:RunLockPath = $null
}

function ConvertTo-MatchText {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Value
    )

    return (($Value.Trim() -replace '\\', '/') -replace '\s*/\s*', '/').ToLowerInvariant()
}

function Get-DiscoveredUnits {
    $units = @()
    foreach ($item in @(Get-ChildItem -LiteralPath $unitsRoot -Directory -Force -ErrorAction Stop)) {
        if ($item.Name -match '^Unit(\d+)$') {
            $units += [pscustomobject] @{
                Name = $item.Name
                Number = [int] $matches[1]
                Path = $item.FullName
            }
        }
    }

    return @($units | Sort-Object Number, Name)
}

function Get-UnitSequenceEntries {
    $sequenceData = Import-PowerShellDataFile -LiteralPath $unitSequencePath
    return @($sequenceData.Workbooks | Sort-Object Sequence | ForEach-Object {
        $folder = [string] $_.Folder
        $fileName = [string] $_.FileName
        $roleName = $null
        if ($folder -match '[\\/]') {
            $roleName = @($folder -split '[\\/]')[-1]
        }

        [pscustomobject] @{
            Sequence = [int] $_.Sequence
            Folder = $folder
            Role = $roleName
            FileName = $fileName
            Stem = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            RelativePath = "$folder/$fileName"
        }
    })
}

function Get-OrgSequenceEntries {
    $sequenceData = Import-PowerShellDataFile -LiteralPath $orgSequencePath
    return @($sequenceData.Workbooks | Sort-Object Sequence | ForEach-Object {
        $relativePath = [string] $_.Path
        $fileName = [System.IO.Path]::GetFileName($relativePath)
        [pscustomobject] @{
            Sequence = [int] $_.Sequence
            RelativePath = $relativePath
            FileName = $fileName
            Stem = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        }
    })
}

function Get-GlobalSequenceEntries {
    $units = @(Get-DiscoveredUnits)
    if ($units.Count -eq 0) {
        throw "No UnitN folders were discovered under $unitsRoot."
    }

    $unitSequenceEntries = @(Get-UnitSequenceEntries)
    $orgSequenceEntries = @(Get-OrgSequenceEntries)
    $globalSequence = 1
    $entries = @()

    foreach ($unit in $units) {
        foreach ($entry in $unitSequenceEntries) {
            $entries += [pscustomobject] @{
                GlobalSequence = $globalSequence
                Stage = "Units"
                UnitName = $unit.Name
                UnitNumber = $unit.Number
                UnitLocalSequence = [int] $entry.Sequence
                OrgLocalSequence = 0
                RelativePath = $entry.RelativePath
                UnitQualifiedPath = ("{0}/{1}" -f $unit.Name, $entry.RelativePath)
                FileName = $entry.FileName
                Stem = $entry.Stem
                Role = $entry.Role
            }
            $globalSequence++
        }
    }

    foreach ($entry in $orgSequenceEntries) {
        $entries += [pscustomobject] @{
            GlobalSequence = $globalSequence
            Stage = "Org"
            UnitName = $null
            UnitNumber = 0
            UnitLocalSequence = 0
            OrgLocalSequence = [int] $entry.Sequence
            RelativePath = $entry.RelativePath
            UnitQualifiedPath = $null
            FileName = $entry.FileName
            Stem = $entry.Stem
            Role = $null
        }
        $globalSequence++
    }

    return $entries
}

function Format-GlobalSequenceEntry {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Entry
    )

    if ($Entry.Stage -eq "Units") {
        return ("{0:000} [Units: {1} #{2:00}] {1} / {3}" -f $Entry.GlobalSequence, $Entry.UnitName, $Entry.UnitLocalSequence, $Entry.RelativePath)
    }

    return ("{0:000} [Org #{1:00}] {2}" -f $Entry.GlobalSequence, $Entry.OrgLocalSequence, $Entry.RelativePath)
}

function Show-GlobalSequence {
    foreach ($entry in @(Get-GlobalSequenceEntries)) {
        Write-Host (Format-GlobalSequenceEntry -Entry $entry)
    }
}

function Get-GlobalEntryMatchTokens {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Entry
    )

    if ($Entry.Stage -eq "Units") {
        $tokens = @(
            $Entry.UnitQualifiedPath,
            ("{0} / {1}" -f $Entry.UnitName, $Entry.RelativePath),
            ("{0} / {1}" -f $Entry.UnitName, $Entry.FileName),
            $Entry.FileName,
            $Entry.Stem
        )

        if (-not [string]::IsNullOrWhiteSpace([string] $Entry.Role)) {
            $tokens += @(
                ("{0} / {1} / {2}" -f $Entry.UnitName, $Entry.Role, $Entry.FileName),
                ("{0} / {1} / {2}" -f $Entry.UnitName, $Entry.Role, $Entry.Stem),
                ("{0} / {1}" -f $Entry.Role, $Entry.FileName),
                ("{0} / {1}" -f $Entry.Role, $Entry.Stem)
            )
        }

        return $tokens
    }

    return @(
        $Entry.RelativePath,
        $Entry.FileName,
        $Entry.Stem
    )
}

function Join-GlobalDisplayNames {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Entries
    )

    return (@($Entries | ForEach-Object { Format-GlobalSequenceEntry -Entry $_ }) -join '; ')
}

function Resolve-GlobalWorkbookStart {
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Entries,

        [Parameter(Mandatory = $true)]
        [string] $WorkbookInput
    )

    $normalizedInput = ConvertTo-MatchText -Value $WorkbookInput
    $exactMatches = @($Entries | Where-Object {
        $entry = $_
        @(Get-GlobalEntryMatchTokens -Entry $entry | ForEach-Object { ConvertTo-MatchText -Value ([string] $_) }) -contains $normalizedInput
    })

    if ($exactMatches.Count -eq 1) {
        return $exactMatches[0]
    }

    if ($exactMatches.Count -gt 1) {
        throw "Start workbook '$WorkbookInput' is ambiguous. Matches: $(Join-GlobalDisplayNames -Entries $exactMatches). Use a global sequence number, unit-qualified path, or org relative path."
    }

    $partialMatches = @($Entries | Where-Object {
        $entry = $_
        $matched = $false
        foreach ($token in @(Get-GlobalEntryMatchTokens -Entry $entry)) {
            if ((ConvertTo-MatchText -Value ([string] $token)).IndexOf($normalizedInput, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $matched = $true
                break
            }
        }
        $matched
    })

    if ($partialMatches.Count -eq 1) {
        return $partialMatches[0]
    }

    if ($partialMatches.Count -gt 1) {
        throw "Start workbook '$WorkbookInput' is ambiguous. Matches: $(Join-GlobalDisplayNames -Entries $partialMatches). Use a global sequence number, unit-qualified path, or org relative path."
    }

    throw "Start workbook '$WorkbookInput' was not found in the Date-level production sequence."
}

function Resolve-GlobalSelection {
    param(
        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    if ($StartSequence -gt 0 -and -not [string]::IsNullOrWhiteSpace($StartWorkbook)) {
        throw "Use StartAtSequence or StartAtWorkbook, not both."
    }

    if ($EndSequence -gt 0 -and -not [string]::IsNullOrWhiteSpace($StartWorkbook)) {
        throw "Use EndAtSequence only with StartAtSequence, not StartAtWorkbook."
    }

    if ($EndSequence -gt 0 -and $StartSequence -lt 1) {
        throw "EndAtSequence requires StartAtSequence."
    }

    if ($EndSequence -gt 0 -and $EndSequence -lt $StartSequence) {
        throw "Invalid global sequence range: end sequence $EndSequence is before start sequence $StartSequence."
    }

    $entries = @(Get-GlobalSequenceEntries)
    if ($entries.Count -eq 0) {
        throw "No production sequence entries were found."
    }

    $totalCount = $entries.Count
    $unitEntriesAll = @($entries | Where-Object { $_.Stage -eq "Units" })
    $orgEntriesAll = @($entries | Where-Object { $_.Stage -eq "Org" })
    $unitTotal = $unitEntriesAll.Count
    $orgTotal = $orgEntriesAll.Count

    if (-not [string]::IsNullOrWhiteSpace($StartWorkbook)) {
        $match = Resolve-GlobalWorkbookStart -Entries $entries -WorkbookInput $StartWorkbook
        $StartSequence = [int] $match.GlobalSequence
    }

    if ($StartSequence -lt 1) {
        $StartSequence = 1
    }

    if ($EndSequence -lt 1) {
        $EndSequence = $totalCount
    }

    if ($StartSequence -gt $totalCount) {
        throw "StartAtSequence $StartSequence is after the end of the Date-level production sequence."
    }

    if ($EndSequence -gt $totalCount) {
        throw "EndAtSequence $EndSequence is after the end of the Date-level production sequence."
    }

    $selectedEntries = @($entries | Where-Object {
        $_.GlobalSequence -ge $StartSequence -and $_.GlobalSequence -le $EndSequence
    })

    if ($selectedEntries.Count -eq 0) {
        throw "Global sequence range $StartSequence-$EndSequence does not match any configured workbook."
    }

    $unitEntries = @($selectedEntries | Where-Object { $_.Stage -eq "Units" })
    $orgEntries = @($selectedEntries | Where-Object { $_.Stage -eq "Org" })

    $unitStart = 0
    $unitEnd = 0
    if ($unitEntries.Count -gt 0) {
        $unitStart = [int] ($unitEntries | Select-Object -ExpandProperty GlobalSequence | Measure-Object -Minimum).Minimum
        $unitEnd = [int] ($unitEntries | Select-Object -ExpandProperty GlobalSequence | Measure-Object -Maximum).Maximum
    }

    $orgStart = 0
    $orgEnd = 0
    if ($orgEntries.Count -gt 0) {
        $orgStart = [int] ($orgEntries | Select-Object -ExpandProperty OrgLocalSequence | Measure-Object -Minimum).Minimum
        $orgEnd = [int] ($orgEntries | Select-Object -ExpandProperty OrgLocalSequence | Measure-Object -Maximum).Maximum
    }

    return [pscustomobject] @{
        Entries = $entries
        SelectedEntries = $selectedEntries
        UnitEntries = $unitEntries
        OrgEntries = $orgEntries
        UnitStart = $unitStart
        UnitEnd = $unitEnd
        OrgStart = $orgStart
        OrgEnd = $orgEnd
        UnitTotal = $unitTotal
        OrgTotal = $orgTotal
        Total = $totalCount
        OrgGlobalStart = $unitTotal + 1
    }
}

function Show-SelectionSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Selection
    )

    Write-Host ("Global production sequence OK: {0} selected workbook(s) of {1} total." -f $Selection.SelectedEntries.Count, $Selection.Total)
    Write-Host ("Unit stage workbook count: {0}" -f $Selection.UnitTotal)
    Write-Host ("Org stage starts at global sequence: {0}" -f $Selection.OrgGlobalStart)
    foreach ($entry in @($Selection.SelectedEntries)) {
        Write-Host (Format-GlobalSequenceEntry -Entry $entry)
    }
}

function Get-OrgWorkflowArguments {
    param(
        [switch] $ValidateOnly,

        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $orgWorkflowScript,
        "-RunRoot", $dateRoot,
        "-SequencePath", $orgSequencePath,
        "-VisibleOverride", $VisibleOverride,
        "-TimeoutMinutesOverride", $TimeoutMinutesOverride
    )

    if ($ValidateOnly) {
        $arguments += "-ValidateSelectionOnly"
    }

    if ($StartSequence -gt 0) {
        $arguments += @("-StartAtSequence", $StartSequence)
    }

    if ($EndSequence -gt 0) {
        $arguments += @("-EndAtSequence", $EndSequence)
    }

    if (-not [string]::IsNullOrWhiteSpace($StartWorkbook)) {
        $arguments += @("-StartAtWorkbook", $StartWorkbook)
    }

    if (-not [string]::IsNullOrWhiteSpace($script:LogFile)) {
        $arguments += @("-LogPath", $script:LogFile)
    }

    if (-not [string]::IsNullOrWhiteSpace($script:StatusFile)) {
        $arguments += @("-StatusPath", $script:StatusFile)
    }

    if (-not [string]::IsNullOrWhiteSpace($stopRequestPath)) {
        $arguments += @("-StopRequestPath", $stopRequestPath)
    }

    return $arguments
}

function Invoke-OrgWorkflow {
    param(
        [switch] $ValidateOnly,

        [int] $StartSequence = 0,

        [int] $EndSequence = 0
    )

    $arguments = Get-OrgWorkflowArguments `
        -ValidateOnly:$ValidateOnly `
        -StartSequence $StartSequence `
        -EndSequence $EndSequence

    Write-DateLog ("Launching integrated org workflow {0} command." -f $(if ($ValidateOnly) { "validation" } else { "refresh" }))
    & pwsh @arguments | ForEach-Object { Write-Host $_ }
    return [int] $LASTEXITCODE
}

function Get-UnitWorkflowArguments {
    param(
        [switch] $ValidateOnly,

        [int] $StartSequence = 0,

        [int] $EndSequence = 0
    )

    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $unitWorkflowScript,
        "-RunRoot", $unitsRoot,
        "-SequencePath", $unitSequencePath,
        "-VisibleOverride", $VisibleOverride,
        "-TimeoutMinutesOverride", $TimeoutMinutesOverride
    )

    if ($ValidateOnly) {
        $arguments += "-ValidateSelectionOnly"
    }

    if ($StartSequence -gt 0) {
        $arguments += @("-StartAtSequence", $StartSequence)
    }

    if ($EndSequence -gt 0) {
        $arguments += @("-EndAtSequence", $EndSequence)
    }

    return $arguments
}

function Invoke-UnitWorkflow {
    param(
        [switch] $ValidateOnly,

        [int] $StartSequence = 0,

        [int] $EndSequence = 0
    )

    $arguments = Get-UnitWorkflowArguments `
        -ValidateOnly:$ValidateOnly `
        -StartSequence $StartSequence `
        -EndSequence $EndSequence

    Write-DateLog ("Launching integrated dynamic all-units {0} command." -f $(if ($ValidateOnly) { "validation" } else { "refresh" }))
    Write-DateLog "Unit workflow script: $unitWorkflowScript"
    & pwsh @arguments | ForEach-Object { Write-Host $_ }
    return [int] $LASTEXITCODE
}

function Show-UnitStatus {
    $statusPath = Join-Path $unitRunLogsPath "current-status.txt"

    Write-Host ""
    Write-Host "Integrated all-units status"
    Write-Host "---------------------------"

    if (Test-Path -LiteralPath $statusPath -PathType Leaf) {
        $statusLines = @(Get-Content -LiteralPath $statusPath)
        $workflowLine = @($statusLines | Where-Object { $_ -match '^workflow=' } | Select-Object -First 1)

        if ($workflowLine.Count -gt 0 -and $workflowLine[0] -ne "workflow=ResidentialCare All Units Refresh") {
            Write-Host "No current ResidentialCare All Units Refresh status file exists yet."
        }
        else {
            $statusLines | ForEach-Object { Write-Host $_ }
        }
    }
    else {
        Write-Host "No unit current-status.txt file exists yet."
    }

    Write-Host ""
    Write-Host "Latest unit-stage log"
    Write-Host "---------------------"

    if (Test-Path -LiteralPath $unitRunLogsPath -PathType Container) {
        $latestLog = Get-ChildItem -LiteralPath $unitRunLogsPath -Filter "AllUnitsRefresh-*.log" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($latestLog) {
            Write-Host $latestLog.FullName
        }
        else {
            Write-Host "No ResidentialCare All Units Refresh log exists yet."
        }

        if (Test-Path -LiteralPath $unitStopRequestPath -PathType Leaf) {
            Write-Host ""
            Write-Host "Unit-stage stop request"
            Write-Host "-----------------------"
            Get-Content -LiteralPath $unitStopRequestPath | ForEach-Object { Write-Host $_ }
        }
    }
    else {
        Write-Host "No unit RunLogs folder exists yet."
    }
}

function Show-Status {
    $statusPath = Join-Path $runLogsPath "current-status.txt"

    Write-Host ""
    Write-Host "ResidentialCare Date Refresh Status"
    Write-Host "-----------------------------------"

    if (Test-Path -LiteralPath $statusPath -PathType Leaf) {
        $statusLines = @(Get-Content -LiteralPath $statusPath)
        $workflowLine = @($statusLines | Where-Object { $_ -match '^workflow=' } | Select-Object -First 1)

        if ($workflowLine.Count -gt 0 -and $workflowLine[0] -ne "workflow=$script:WorkflowName") {
            Write-Host "No current ResidentialCare Date Refresh status file exists yet."
        }
        else {
            $statusLines | ForEach-Object { Write-Host $_ }
        }
    }
    else {
        Write-Host "No current-status.txt file exists yet."
    }

    Write-Host ""
    Write-Host "Latest Date-level log"
    Write-Host "---------------------"

    if (Test-Path -LiteralPath $runLogsPath -PathType Container) {
        $latestLog = Get-ChildItem -LiteralPath $runLogsPath -Filter "ResidentialCareRefresh-*.log" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($latestLog) {
            Write-Host $latestLog.FullName
        }
        else {
            Write-Host "No ResidentialCare Date Refresh log exists yet."
        }

        if (Test-Path -LiteralPath $stopRequestPath -PathType Leaf) {
            Write-Host ""
            Write-Host "Date-level stop request"
            Write-Host "-----------------------"
            Get-Content -LiteralPath $stopRequestPath | ForEach-Object { Write-Host $_ }
        }
    }
    else {
        Write-Host "No Date-level RunLogs folder exists yet."
    }

    Show-UnitStatus
}

function Write-UnitStopRequest {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("AfterCurrent", "Now")]
        [string] $Mode
    )

    if (-not (Test-Path -LiteralPath $unitRunLogsPath -PathType Container)) {
        New-Item -ItemType Directory -Path $unitRunLogsPath -Force | Out-Null
    }

    $lines = @(
        "mode=$Mode",
        "requestedAt=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "source=ResidentialCare Date Refresh integrated unit stage"
    )

    Set-Content -LiteralPath $unitStopRequestPath -Value $lines
    Write-Host "Unit-stage stop request written: $unitStopRequestPath"
}

function Write-StopRequest {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("AfterCurrent", "Now")]
        [string] $Mode
    )

    if (-not (Test-Path -LiteralPath $runLogsPath -PathType Container)) {
        New-Item -ItemType Directory -Path $runLogsPath -Force | Out-Null
    }

    $lines = @(
        "mode=$Mode",
        "requestedAt=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "source=ResidentialCare Date Refresh"
    )

    Set-Content -LiteralPath $stopRequestPath -Value $lines
    Write-Host "Date-level stop request written: $stopRequestPath"
    Write-Host "Mode: $Mode"

    Write-Host ""
    Write-UnitStopRequest -Mode $Mode
}

function Invoke-Validation {
    param(
        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    Write-Host "Validating ResidentialCare Date Refresh."
    $selection = Resolve-GlobalSelection -StartSequence $StartSequence -EndSequence $EndSequence -StartWorkbook $StartWorkbook
    Show-SelectionSummary -Selection $selection

    if ($selection.UnitEntries.Count -gt 0) {
        Write-Host ""
        Write-Host "Step 1: integrated dynamic all-units validation."
        $unitValidationExitCode = Invoke-UnitWorkflow -ValidateOnly -StartSequence $selection.UnitStart -EndSequence $selection.UnitEnd
        if ($unitValidationExitCode -ne 0) {
            Write-Host "Unit validation failed with exit code $unitValidationExitCode."
            return $unitValidationExitCode
        }
    }
    else {
        Write-Host ""
        Write-Host "Step 1: integrated dynamic all-units validation skipped because selection starts at org stage."
    }

    if ($selection.OrgEntries.Count -gt 0) {
        Write-Host ""
        Write-Host "Step 2: Date-level org calculation validation."
        $orgValidationExitCode = Invoke-OrgWorkflow -ValidateOnly -StartSequence $selection.OrgStart -EndSequence $selection.OrgEnd
        if ($orgValidationExitCode -ne 0) {
            Write-Host "Org validation failed with exit code $orgValidationExitCode."
            return $orgValidationExitCode
        }
    }
    else {
        Write-Host ""
        Write-Host "Step 2: Date-level org calculation validation skipped because selection ends before org stage."
    }

    Write-Host ""
    Write-Host "ResidentialCare Date Refresh validation passed."
    return 0
}

function Invoke-RefreshSelection {
    param(
        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    if (-not (Test-Path -LiteralPath $runLogsPath -PathType Container)) {
        New-Item -ItemType Directory -Path $runLogsPath -Force | Out-Null
    }

    $selection = Resolve-GlobalSelection -StartSequence $StartSequence -EndSequence $EndSequence -StartWorkbook $StartWorkbook
    $lockPath = Join-Path $runLogsPath "ResidentialCareRefresh.lock"
    Acquire-DateRunLock -LockPath $lockPath

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $script:LogFile = Join-Path $runLogsPath ("ResidentialCareRefresh-{0}.log" -f $timestamp)
    $script:StatusFile = Join-Path $runLogsPath "current-status.txt"

    try {
        if (Test-Path -LiteralPath $stopRequestPath -PathType Leaf) {
            Remove-Item -LiteralPath $stopRequestPath -Force
        }

        Write-DateStatus -State "RUNNING" -Stage "Preflight"
        Write-DateLog "Starting ResidentialCare Date Refresh."
        Write-DateLog "Date root: $dateRoot"
        Write-DateLog "Date-level lock: $lockPath"
        Write-DateLog "Date-level log: $script:LogFile"
        Write-DateLog ("Selected global sequence range: {0}-{1}" -f $selection.SelectedEntries[0].GlobalSequence, $selection.SelectedEntries[-1].GlobalSequence)

        if ($selection.UnitEntries.Count -gt 0) {
            Write-DateStatus -State "RUNNING" -Stage "Unit stage"
            Write-DateLog "Unit stage start."
            $unitRefreshExitCode = Invoke-UnitWorkflow -StartSequence $selection.UnitStart -EndSequence $selection.UnitEnd
            if ($unitRefreshExitCode -ne 0) {
                $script:ExitCode = $unitRefreshExitCode
                throw "Unit stage failed with exit code $unitRefreshExitCode. Org stage was not started."
            }
            Write-DateLog "Unit stage succeeded."
        }
        else {
            Write-DateLog "Unit stage skipped because selection starts at org stage."
        }

        if ($selection.OrgEntries.Count -gt 0) {
            Write-DateStatus -State "RUNNING" -Stage "Org stage"
            Write-DateLog "Org stage start."
            $orgRefreshExitCode = Invoke-OrgWorkflow -StartSequence $selection.OrgStart -EndSequence $selection.OrgEnd
            if ($orgRefreshExitCode -ne 0) {
                $script:ExitCode = $orgRefreshExitCode
                throw "Org stage failed with exit code $orgRefreshExitCode."
            }
            Write-DateLog "Org stage succeeded."
        }
        else {
            Write-DateLog "Org stage skipped because selection ends before org stage."
        }

        Write-DateLog "ResidentialCare Date Refresh succeeded."
        Write-DateStatus -State "SUCCEEDED" -Message "ResidentialCare Date Refresh completed."
        $script:ExitCode = 0
        return 0
    }
    catch {
        $message = $_.Exception.Message
        if ($script:ExitCode -eq 0) {
            $script:ExitCode = 1
        }

        Write-DateLog $message "ERROR"
        Write-DateStatus -State "FAILED" -Message $message
        return $script:ExitCode
    }
    finally {
        Release-DateRunLock
    }
}

function Invoke-ValidatedRefresh {
    param(
        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    $validationExitCode = Invoke-Validation -StartSequence $StartSequence -EndSequence $EndSequence -StartWorkbook $StartWorkbook
    if ($validationExitCode -ne 0) {
        Write-Host "Validation failed with exit code $validationExitCode. Excel was not started."
        return $validationExitCode
    }

    Write-Host ""
    Write-Host "Validation passed. Starting refresh."
    return (Invoke-RefreshSelection -StartSequence $StartSequence -EndSequence $EndSequence -StartWorkbook $StartWorkbook)
}

function Get-UnitWorkbookCount {
    return @(Get-UnitSequenceEntries).Count
}

function Show-UnitsMenu {
    $units = @(Get-DiscoveredUnits)
    if ($units.Count -eq 0) {
        Write-Host "No UnitN folders were discovered."
        return
    }

    $perUnitWorkbookCount = Get-UnitWorkbookCount
    while ($true) {
        Write-Host ""
        Write-Host "Refresh UNITS"
        Write-Host "============="
        Write-Host "1. All discovered units"

        for ($i = 0; $i -lt $units.Count; $i++) {
            Write-Host ("{0}. {1}" -f ($i + 2), $units[$i].Name)
        }

        $backOption = $units.Count + 2
        Write-Host ("{0}. Back" -f $backOption)
        Write-Host ""

        $choiceText = Read-Host "Choose a unit option"
        $choice = 0
        if (-not [int]::TryParse($choiceText, [ref] $choice)) {
            Write-Host "Choose a menu number."
            continue
        }

        if ($choice -eq 1) {
            $unitTotal = $units.Count * $perUnitWorkbookCount
            Invoke-ValidatedRefresh -StartSequence 1 -EndSequence $unitTotal | Out-Null
            continue
        }

        if ($choice -eq $backOption) {
            return
        }

        $unitIndex = $choice - 2
        if ($unitIndex -lt 0 -or $unitIndex -ge $units.Count) {
            Write-Host "Choose a listed unit option."
            continue
        }

        $startSequence = ($unitIndex * $perUnitWorkbookCount) + 1
        $endSequence = $startSequence + $perUnitWorkbookCount - 1
        Invoke-ValidatedRefresh -StartSequence $startSequence -EndSequence $endSequence | Out-Null
    }
}

function Read-SequenceNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Prompt
    )

    $value = Read-Host $Prompt
    $number = 0
    if (-not [int]::TryParse($value, [ref] $number) -or $number -lt 1) {
        Write-Host "Enter a positive global sequence number."
        return $null
    }

    return $number
}

function Show-Menu {
    while ($true) {
        Write-Host ""
        Write-Host "ResidentialCare Date Refresh"
        Write-Host "============================"
        Write-Host "1. Refresh production flow"
        Write-Host "2. Validate production flow"
        Write-Host "3. Refresh UNITS"
        Write-Host "4. Refresh Org level"
        Write-Host "5. Refresh from workbook name"
        Write-Host "6. Refresh from global sequence"
        Write-Host "7. Stop after current workbook"
        Write-Host "8. Stop now"
        Write-Host "9. Show current status / latest logs"
        Write-Host "10. Exit"
        Write-Host ""

        $choice = Read-Host "Choose an option"

        switch ($choice) {
            "1" {
                Write-Host "Selected option 1: Refresh production flow."
                Write-Host "First it runs all discovered units, then refreshes the 8 org calculation workbooks in order."
                Invoke-ValidatedRefresh | Out-Null
            }
            "2" {
                Write-Host "Selected option 2: Validate production flow."
                Invoke-Validation | Out-Null
            }
            "3" {
                Write-Host "Selected option 3: Refresh UNITS."
                Show-UnitsMenu
            }
            "4" {
                Write-Host "Selected option 4: Refresh Org level."
                $selection = Resolve-GlobalSelection
                Invoke-ValidatedRefresh -StartSequence $selection.OrgGlobalStart | Out-Null
            }
            "5" {
                Write-Host "Selected option 5: Refresh from workbook name."
                Write-Host ""
                Show-GlobalSequence
                Write-Host ""
                $workbook = Read-Host "Enter workbook name, unit-qualified path, or org relative path"
                if ([string]::IsNullOrWhiteSpace($workbook)) {
                    Write-Host "No workbook entered. Excel was not started."
                    continue
                }

                Invoke-ValidatedRefresh -StartWorkbook $workbook | Out-Null
            }
            "6" {
                Write-Host "Selected option 6: Refresh from global sequence."
                Write-Host ""
                Show-GlobalSequence
                Write-Host ""
                $sequence = Read-SequenceNumber -Prompt "Enter start global sequence number"
                if ($null -eq $sequence) {
                    Write-Host "Excel was not started."
                    continue
                }

                $endText = Read-Host "Optional end global sequence number, or press Enter to run to the end"
                if ([string]::IsNullOrWhiteSpace($endText)) {
                    Invoke-ValidatedRefresh -StartSequence $sequence | Out-Null
                    continue
                }

                $end = 0
                if (-not [int]::TryParse($endText, [ref] $end) -or $end -lt 1) {
                    Write-Host "End sequence must be a positive global sequence number. Excel was not started."
                    continue
                }

                Invoke-ValidatedRefresh -StartSequence $sequence -EndSequence $end | Out-Null
            }
            "7" {
                Write-Host "Selected option 7: Stop after current workbook."
                Write-StopRequest -Mode "AfterCurrent"
            }
            "8" {
                Write-Host "Selected option 8: Stop now."
                Write-StopRequest -Mode "Now"
            }
            "9" {
                Write-Host "Selected option 9: Show current status / latest logs."
                Show-Status
            }
            "10" {
                Write-Host "Exiting. Excel was not started."
                return
            }
            default {
                Write-Host "Choose a menu number from 1 to 10."
            }
        }
    }
}

try {
    Assert-PackageFolder -Path $unitsRoot -Description "UNITS root"
    Assert-PackageFile -Path $unitWorkflowScript -Description "integrated all-units workflow script"
    Assert-PackageFile -Path $unitSequencePath -Description "integrated all-units workbook sequence file"
    Assert-PackageFile -Path $orgWorkflowScript -Description "Date-level org workflow script"
    Assert-PackageFile -Path $orgSequencePath -Description "Date-level org workbook sequence file"

    if ($ShowStatus) {
        Show-Status
        exit 0
    }

    if (-not [string]::IsNullOrWhiteSpace($StopMode)) {
        Write-StopRequest -Mode $StopMode
        exit 0
    }

    if ($ValidateSelectionOnly) {
        $exitCode = Invoke-Validation -StartSequence $StartAtSequence -EndSequence $EndAtSequence -StartWorkbook $StartAtWorkbook
        exit $exitCode
    }

    $hasExplicitSelection =
        $RefreshAll -or
        $StartAtSequence -gt 0 -or
        $EndAtSequence -gt 0 -or
        -not [string]::IsNullOrWhiteSpace($StartAtWorkbook)

    if ($hasExplicitSelection) {
        $exitCode = Invoke-ValidatedRefresh -StartSequence $StartAtSequence -EndSequence $EndAtSequence -StartWorkbook $StartAtWorkbook
        exit $exitCode
    }

    Show-Menu
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
