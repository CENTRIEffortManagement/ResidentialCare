# Legacy coordinators share read/write leases; batch workers share read-only leases
# behind a single-coordinator lock. Nested legacy stages can coexist, but cannot
# overlap a batch run. Never
# delete this file: unlinking a lock introduces an acquisition/deletion race.
function Find-RefreshDateRoot {
    param([string] $Path)
    $candidate = [IO.Path]::GetFullPath($Path)
    while ($candidate) {
        if (Test-Path -LiteralPath (Join-Path $candidate 'runner/ResidentialCare-ClientRunProfile.psd1')) { return $candidate }
        $candidate = Split-Path $candidate -Parent
    }
    throw 'Cannot locate the ResidentialCare Date root.'
}

function Enter-RefreshRunGate {
    param([string] $DateRoot, [switch] $Exclusive, [switch] $BatchWorker)
    $folder = Join-Path $DateRoot 'RunLogs'
    [void] [IO.Directory]::CreateDirectory($folder)
    $path = Join-Path $folder 'BatchRefresh.gate'
    # Batch workers retain a read-only family lease if their coordinator dies.
    # Legacy stages request write access and cannot overlap those leases.
    $access = if ($Exclusive -or $BatchWorker) { [IO.FileAccess]::Read } else { [IO.FileAccess]::ReadWrite }
    $share = if ($Exclusive -or $BatchWorker) { [IO.FileShare]::Read } else { [IO.FileShare]::ReadWrite }
    $coordinator = $null
    try {
        # CreateNew is only for initialization, not lock acquisition.
        if (-not [IO.File]::Exists($path)) {
            try { [IO.File]::Open($path, 'CreateNew', 'ReadWrite', 'ReadWrite').Dispose() }
            catch [IO.IOException] { if (-not [IO.File]::Exists($path)) { throw } }
        }
        if ($Exclusive) {
            $coordinator = [IO.File]::Open((Join-Path $folder 'BatchCoordinator.gate'), 'OpenOrCreate', 'ReadWrite', 'None')
            # Detect orphan workers as well as legacy runs before admitting a new run.
            [IO.File]::Open($path, 'Open', 'ReadWrite', 'None').Dispose()
        }
        $lease = [pscustomobject]@{ Family = [IO.File]::Open($path, 'Open', $access, $share); Coordinator = $coordinator }
        $lease | Add-Member ScriptMethod Dispose {
            $this.Family.Dispose()
            if ($null -ne $this.Coordinator) { $this.Coordinator.Dispose() }
        }
        return $lease
    }
    catch {
        if ($null -ne $coordinator) { $coordinator.Dispose() }
        throw "Conflicting ResidentialCare runner is active (Date-level gate): $path. $($_.Exception.Message)"
    }
}

function Assert-RefreshRunGateAvailable {
    param([string] $DateRoot)
    $path = Join-Path $DateRoot 'RunLogs/BatchRefresh.gate'
    if ([IO.File]::Exists($path)) {
        try { [IO.File]::Open($path, 'Open', 'ReadWrite', 'None').Dispose() }
        catch { throw 'A conflicting ResidentialCare runner is active.' }
    }
}
