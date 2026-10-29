[CmdletBinding()]
param([string] $ProjectRoot = (Get-Location).Path)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath $ProjectRoot).ProviderPath.TrimEnd('\')
$repoPrefix = $repoRoot + '\'
$runRoot = (Resolve-Path -LiteralPath $PSScriptRoot).ProviderPath

function Assert-WorkspacePath([string] $Path) {
    $absolute = [IO.Path]::GetFullPath($Path)
    if (-not $absolute.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside the repository: $absolute"
    }
    $current = $absolute
    while (-not $current.Equals($repoRoot, [StringComparison]::OrdinalIgnoreCase)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Reparse point is outside this operation: $current"
            }
        }
        $current = [IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrEmpty($current)) { throw 'Path ancestry validation failed' }
    }
    return $absolute
}

function Relative-Path([string] $Path) {
    return [IO.Path]::GetRelativePath($repoRoot, $Path).Replace('\', '/')
}

function Assert-Unlocked([string] $Path) {
    $handle = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    $handle.Dispose()
}

$null = Assert-WorkspacePath $runRoot
$config = Get-Content -LiteralPath (Join-Path $repoRoot 'pq.project.json') -Raw | ConvertFrom-Json
$approvalPath = Assert-WorkspacePath (Join-Path $repoRoot $config.batchRunner.approval)
$approval = Import-PowerShellDataFile -LiteralPath $approvalPath
$units = @($approval.AdditionalUnits)
if ((@($units.Folder | Sort-Object) -join ',') -ne 'BD,JH-RY,TE') { throw 'Unexpected production unit scope' }

$donorRoot = Assert-WorkspacePath (Join-Path $repoRoot ($config.batchRunner.dateRoot + '/UNITS/BD/2. Calculations/AIN'))
$names = @('CapacityDistrib(A.1)-shifts.xlsx', 'CapacityDistrib(A.2)-shifts.xlsx', 'CapacityDistrib(B)-shifts.xlsx')
$donors = @()
foreach ($name in $names) {
    $path = Assert-WorkspacePath (Join-Path $donorRoot $name)
    $item = Get-Item -LiteralPath $path
    if ($item.PSIsContainer) { throw 'Donor is not a file' }
    Assert-Unlocked $path
    $donors += [pscustomobject]@{ Name = $name; Path = $path; SHA256 = (Get-FileHash -LiteralPath $path).Hash; Length = $item.Length }
}

$targets = @()
$originals = @()
foreach ($unit in $units) {
    if ((@($unit.Roles | Sort-Object) -join ',') -ne 'AIN,CS,EN,RN') { throw 'Unexpected role scope' }
    foreach ($role in $unit.Roles) {
        $path = Assert-WorkspacePath (Join-Path $repoRoot ($config.batchRunner.dateRoot + '/UNITS/' + $unit.Folder + '/2. Calculations/' + $role))
        $resolved = (Resolve-Path -LiteralPath $path).ProviderPath
        if (-not $resolved.Equals($path, [StringComparison]::OrdinalIgnoreCase)) { throw 'Target path resolution changed' }
        if ($path.Equals($donorRoot, [StringComparison]::OrdinalIgnoreCase)) { continue }
        $items = @(Get-ChildItem -LiteralPath $path -Force)
        if (@($items | Where-Object PSIsContainer).Count -gt 0) { throw "Unexpected subdirectory: $path" }
        $targets += [pscustomobject]@{ Unit = $unit.Folder; Role = $role; Path = $path }
        foreach ($item in $items) {
            $filePath = Assert-WorkspacePath $item.FullName
            Assert-Unlocked $filePath
            $originals += [pscustomobject]@{
                Unit = $unit.Folder; Role = $role; Path = $filePath; Name = $item.Name
                SHA256 = (Get-FileHash -LiteralPath $filePath).Hash; Length = $item.Length
                Backup = ''; Deleted = $false
            }
        }
    }
}
if ($targets.Count -ne 11) { throw 'Expected exactly 11 destinations' }

$manifestPath = Join-Path $runRoot 'manifest.json'
$manifest = [ordered]@{
    Operation = 'User-requested clearing of destination role files and exact BD/AIN Excel copies'
    CreatedSydney = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTimeOffset]::UtcNow, 'AUS Eastern Standard Time').ToString('o')
    Status = 'Preparing'; Donor = (Relative-Path $donorRoot); DonorPreserved = $true
    TargetFolders = @($targets | ForEach-Object { Relative-Path $_.Path })
    Sources = @($donors | ForEach-Object { [ordered]@{ Path = (Relative-Path $_.Path); SHA256 = $_.SHA256; Length = $_.Length } })
    RemovedFiles = @(); Copies = @(); SyncPerformed = $false; RefreshPerformed = $false; Error = $null
}
function Save-Manifest { $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding utf8 }
Save-Manifest

try {
    # Verify recoverable copies of every destination file before deleting anything.
    foreach ($file in $originals) {
        $backupFolder = Assert-WorkspacePath (Join-Path $runRoot ('backups/' + $file.Unit + '/' + $file.Role))
        $null = New-Item -ItemType Directory -Path $backupFolder -Force
        $backupPath = Assert-WorkspacePath (Join-Path $backupFolder $file.Name)
        Copy-Item -LiteralPath $file.Path -Destination $backupPath -Force
        if ((Get-FileHash -LiteralPath $backupPath).Hash -cne $file.SHA256) { throw 'Backup verification failed' }
        $file.Backup = Relative-Path $backupPath
    }
    $staging = Assert-WorkspacePath (Join-Path $runRoot 'donor')
    $null = New-Item -ItemType Directory -Path $staging
    foreach ($donor in $donors) {
        $stagedPath = Join-Path $staging $donor.Name
        Copy-Item -LiteralPath $donor.Path -Destination $stagedPath
        if ((Get-FileHash -LiteralPath $stagedPath).Hash -cne $donor.SHA256) { throw 'Donor staging verification failed' }
    }
    # Recheck the exact target paths, complete file lists, hashes and locks immediately before deletion.
    foreach ($target in $targets) {
        $null = Assert-WorkspacePath $target.Path
        $expected = @($originals | Where-Object { $_.Unit -eq $target.Unit -and $_.Role -eq $target.Role })
        $actual = @(Get-ChildItem -LiteralPath $target.Path -Force)
        if ($actual.Count -ne $expected.Count -or @($actual | Where-Object PSIsContainer).Count -gt 0) { throw 'Target contents changed during preflight' }
        foreach ($file in $expected) {
            Assert-Unlocked $file.Path
            if ((Get-FileHash -LiteralPath $file.Path).Hash -cne $file.SHA256) { throw 'Target changed during backup' }
        }
    }
    foreach ($donor in $donors) {
        Assert-Unlocked $donor.Path
        if ((Get-FileHash -LiteralPath $donor.Path).Hash -cne $donor.SHA256) { throw 'Donor changed during preflight' }
    }
    $manifest.Status = 'BackupsVerified'
    $manifest.RemovedFiles = @($originals | ForEach-Object { [ordered]@{ Path = (Relative-Path $_.Path); Backup = $_.Backup; SHA256 = $_.SHA256; Deleted = $false } })
    Save-Manifest

    foreach ($target in $targets) {
        foreach ($file in @($originals | Where-Object { $_.Unit -eq $target.Unit -and $_.Role -eq $target.Role })) {
            $null = Assert-WorkspacePath $file.Path
            if (-not ([IO.Path]::GetDirectoryName($file.Path)).Equals($target.Path, [StringComparison]::OrdinalIgnoreCase)) { throw 'Delete target is not a direct role file' }
            Remove-Item -LiteralPath $file.Path -Force
            $file.Deleted = $true
            ($manifest.RemovedFiles | Where-Object Path -eq (Relative-Path $file.Path)).Deleted = $true
        }
        foreach ($donor in $donors) {
            $destination = Assert-WorkspacePath (Join-Path $target.Path $donor.Name)
            Copy-Item -LiteralPath (Join-Path $staging $donor.Name) -Destination $destination
            $hash = (Get-FileHash -LiteralPath $destination).Hash
            if ($hash -cne $donor.SHA256) { throw 'Destination copy verification failed' }
            $manifest.Copies += [ordered]@{ Source = (Relative-Path $donor.Path); Destination = (Relative-Path $destination); SHA256 = $hash; Verified = $true }
        }
        $actualNames = @(Get-ChildItem -LiteralPath $target.Path -Force | Select-Object -ExpandProperty Name | Sort-Object)
        if (($actualNames -join '|') -cne (($names | Sort-Object) -join '|')) { throw 'Final role contents do not match the three Excel files' }
        Save-Manifest
    }
    foreach ($donor in $donors) {
        if ((Get-FileHash -LiteralPath $donor.Path).Hash -cne $donor.SHA256) { throw 'Donor changed during operation' }
    }
    $manifest.Status = 'Complete'
    Save-Manifest
    [ordered]@{ Status = 'Complete'; DestinationFolders = $targets.Count; DeletedFiles = @($originals | Where-Object Deleted).Count; VerifiedExcelCopies = $manifest.Copies.Count; DonorPreserved = $true; Manifest = (Relative-Path $manifestPath); SyncPerformed = $false; RefreshPerformed = $false } | ConvertTo-Json
}
catch {
    $manifest.Status = 'Failed'
    $manifest.Error = $_.Exception.Message
    Save-Manifest
    throw
}
