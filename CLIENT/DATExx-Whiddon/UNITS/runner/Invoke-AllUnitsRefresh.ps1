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

$runnerRoot = $PSScriptRoot
$runRoot = Split-Path -Path $runnerRoot -Parent
$workflowScript = Join-Path $runnerRoot "src\Invoke-WorkflowRefresh.ps1"
$sequencePath = Join-Path $runnerRoot "AllUnits-WorkbookSequence.psd1"
$runLogsPath = Join-Path $runRoot "RunLogs"
$stopRequestPath = Join-Path $runLogsPath "stop-request.txt"

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

function Get-WorkflowArguments {
    param(
        [switch] $ValidateOnly,

        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    $arguments = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $workflowScript,
        "-RunRoot", $runRoot,
        "-SequencePath", $sequencePath,
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

    return $arguments
}

function Invoke-Workflow {
    param(
        [switch] $ValidateOnly,

        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    $arguments = Get-WorkflowArguments `
        -ValidateOnly:$ValidateOnly `
        -StartSequence $StartSequence `
        -EndSequence $EndSequence `
        -StartWorkbook $StartWorkbook

    Write-Host ("Launching all-units workflow {0} command..." -f $(if ($ValidateOnly) { "validation" } else { "refresh" }))
    Write-Host ("Workflow script: {0}" -f $workflowScript)
    & pwsh @arguments | ForEach-Object { Write-Host $_ }
    $exitCode = [int] $LASTEXITCODE
    return $exitCode
}

function Get-DiscoveredUnits {
    $units = @()
    foreach ($item in @(Get-ChildItem -LiteralPath $runRoot -Directory -Force -ErrorAction Stop)) {
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

function Get-WorkbookDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Entry
    )

    $folder = [string] $Entry.Folder
    $fileName = [string] $Entry.FileName
    if ($folder -match '[\\/]') {
        $roleName = @($folder -split '[\\/]')[-1]
        return "$roleName / $fileName"
    }

    return "$folder/$fileName"
}

function Show-Sequence {
    try {
        $units = @(Get-DiscoveredUnits)
        $data = Import-PowerShellDataFile -LiteralPath $sequencePath
        $workbooks = @($data.Workbooks | Sort-Object Sequence)
        $globalSequence = 1

        if ($units.Count -eq 0) {
            Write-Host "No UnitN folders were discovered under $runRoot."
            return
        }

        Write-Host ("Discovered units: {0}" -f (($units | ForEach-Object { $_.Name }) -join ', '))
        foreach ($unit in $units) {
            foreach ($entry in $workbooks) {
                Write-Host ("{0:000} [{1} #{2:00}] {1} / {3}" -f $globalSequence, $unit.Name, [int] $entry.Sequence, (Get-WorkbookDisplayName -Entry $entry))
                $globalSequence++
            }
        }
    }
    catch {
        Write-Host "Could not read all-units workbook sequence: $($_.Exception.Message)"
    }
}

function Show-Status {
    $statusPath = Join-Path $runLogsPath "current-status.txt"

    Write-Host ""
    Write-Host "Status"
    Write-Host "------"

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
        Write-Host "No current-status.txt file exists yet."
    }

    Write-Host ""
    Write-Host "Latest log"
    Write-Host "----------"

    if (Test-Path -LiteralPath $runLogsPath -PathType Container) {
        $latestLog = Get-ChildItem -LiteralPath $runLogsPath -Filter "AllUnitsRefresh-*.log" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($latestLog) {
            Write-Host $latestLog.FullName
        }
        else {
            Write-Host "No ResidentialCare All Units Refresh log exists yet."
        }

        if (Test-Path -LiteralPath $stopRequestPath -PathType Leaf) {
            Write-Host ""
            Write-Host "Stop request"
            Write-Host "------------"
            Get-Content -LiteralPath $stopRequestPath | ForEach-Object { Write-Host $_ }
        }
    }
    else {
        Write-Host "No RunLogs folder exists yet."
    }
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
        "source=ResidentialCare All Units Refresh menu"
    )

    Set-Content -LiteralPath $stopRequestPath -Value $lines
    Write-Host "Stop request written: $stopRequestPath"
    Write-Host "Mode: $Mode"
}

function Confirm-AndRunRefresh {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Description,

        [int] $StartSequence = 0,

        [int] $EndSequence = 0,

        [string] $StartWorkbook = $null
    )

    Write-Host ""
    Write-Host "Menu selection: $Description"
    Write-Host "Checking selection before Excel starts..."
    $validationExitCode = Invoke-Workflow `
        -ValidateOnly `
        -StartSequence $StartSequence `
        -EndSequence $EndSequence `
        -StartWorkbook $StartWorkbook

    if ($validationExitCode -ne 0) {
        Write-Host "Validation failed with exit code $validationExitCode. Excel was not started."
        return
    }

    Write-Host ""
    Write-Host "Validation passed."
    Write-Host "Starting refresh command: $Description"

    $refreshExitCode = Invoke-Workflow `
        -StartSequence $StartSequence `
        -EndSequence $EndSequence `
        -StartWorkbook $StartWorkbook

    Write-Host "Refresh command finished with exit code $refreshExitCode."
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
        Write-Host "ResidentialCare All Units Refresh"
        Write-Host "================================="
        Write-Host "1. Refresh all discovered unit workbooks"
        Write-Host "2. Refresh from workbook"
        Write-Host "3. Refresh from global sequence number"
        Write-Host "4. Refresh global sequence range"
        Write-Host "5. Stop after current workbook"
        Write-Host "6. Stop now"
        Write-Host "7. Show current status / latest log"
        Write-Host "8. Exit"
        Write-Host ""

        $choice = Read-Host "Choose an option"

        switch ($choice) {
            "1" {
                Write-Host "Selected option 1: Refresh all discovered unit workbooks."
                Confirm-AndRunRefresh -Description "all configured workbooks in all discovered units"
            }
            "2" {
                Write-Host "Selected option 2: Refresh from workbook."
                Write-Host ""
                Show-Sequence
                $workbook = Read-Host "Enter unit-qualified workbook name, exact relative path, or partial name"
                if ([string]::IsNullOrWhiteSpace($workbook)) {
                    Write-Host "No workbook entered. Excel was not started."
                    continue
                }

                Confirm-AndRunRefresh -Description "from workbook '$workbook'" -StartWorkbook $workbook
            }
            "3" {
                Write-Host "Selected option 3: Refresh from global sequence number."
                Write-Host ""
                Show-Sequence
                $sequence = Read-SequenceNumber -Prompt "Enter start global sequence number"
                if ($null -eq $sequence) {
                    Write-Host "Excel was not started."
                    continue
                }

                Confirm-AndRunRefresh -Description "from global sequence $sequence" -StartSequence $sequence
            }
            "4" {
                Write-Host "Selected option 4: Refresh global sequence range."
                Write-Host ""
                Show-Sequence
                $rangeText = Read-Host "Enter inclusive global sequence range, for example 1-26 or 27-52"
                if ($rangeText -notmatch '^\s*(\d+)\s*-\s*(\d+)\s*$') {
                    Write-Host "Range must use the format start-end, such as 27-52. Excel was not started."
                    continue
                }

                $start = [int] $matches[1]
                $end = [int] $matches[2]
                Confirm-AndRunRefresh -Description "global sequence range $start-$end" -StartSequence $start -EndSequence $end
            }
            "5" {
                Write-Host "Selected option 5: Stop after current workbook."
                Write-StopRequest -Mode "AfterCurrent"
            }
            "6" {
                Write-Host "Selected option 6: Stop now."
                Write-StopRequest -Mode "Now"
            }
            "7" {
                Write-Host "Selected option 7: Show current status / latest log."
                Show-Status
            }
            "8" {
                Write-Host "Exiting. Excel was not started."
                return
            }
            default {
                Write-Host "Choose a menu number from 1 to 8."
            }
        }
    }
}

try {
    Assert-PackageFile -Path $workflowScript -Description "workflow script"
    Assert-PackageFile -Path $sequencePath -Description "all-units workbook sequence file"

    if ($ShowStatus) {
        Show-Status
        exit 0
    }

    if (-not [string]::IsNullOrWhiteSpace($StopMode)) {
        Write-StopRequest -Mode $StopMode
        exit 0
    }

    $hasExplicitSelection =
        $RefreshAll -or
        $StartAtSequence -gt 0 -or
        $EndAtSequence -gt 0 -or
        -not [string]::IsNullOrWhiteSpace($StartAtWorkbook)

    if ($ValidateSelectionOnly) {
        $exitCode = Invoke-Workflow `
            -ValidateOnly `
            -StartSequence $StartAtSequence `
            -EndSequence $EndAtSequence `
            -StartWorkbook $StartAtWorkbook

        exit $exitCode
    }

    if ($hasExplicitSelection) {
        $validationExitCode = Invoke-Workflow `
            -ValidateOnly `
            -StartSequence $StartAtSequence `
            -EndSequence $EndAtSequence `
            -StartWorkbook $StartAtWorkbook

        if ($validationExitCode -ne 0) {
            exit $validationExitCode
        }

        $exitCode = Invoke-Workflow `
            -StartSequence $StartAtSequence `
            -EndSequence $EndAtSequence `
            -StartWorkbook $StartAtWorkbook

        exit $exitCode
    }

    Show-Menu
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
