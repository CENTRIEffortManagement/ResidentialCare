Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'RunScope.ps1')

function Resolve-BatchPath {
    param([string] $Root, [string] $RelativePath)
    if ([IO.Path]::IsPathRooted($RelativePath)) { throw "Expected a relative path: $RelativePath" }
    $base = [IO.Path]::GetFullPath($Root).TrimEnd([char[]] '\/')
    $path = [IO.Path]::GetFullPath((Join-Path $base $RelativePath))
    if (-not $path.StartsWith($base + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path escapes its approved root: $RelativePath"
    }
    # Reject aliases/junctions so two job names cannot secretly write the same file.
    $part = $path
    while ($part.Length -ge $base.Length) {
        if (Test-Path -LiteralPath $part) {
            if ((Get-Item -LiteralPath $part -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Reparse-point paths require an explicit canonical configuration: $RelativePath"
            }
        }
        $part = Split-Path $part -Parent
    }
    return $path
}

function Get-BatchFingerprint {
    param([string[]] $Files)
    $parts = foreach ($file in $Files) {
        # Configuration identity is retained for safe resume, not live approval.
        [IO.File]::ReadAllText($file)
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($parts -join "`n---`n"))
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

function Expand-BatchArguments {
    param([string[]] $Values)
    return @($Values | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
}

function Get-BatchCatalogue {
    param([string] $RepoRoot,
        [ValidateSet('Current', 'ParallelInputs-Unit1Priority')] [string] $SequenceProfile = 'Current')
    $project = Get-Content -LiteralPath (Join-Path $RepoRoot 'pq.project.json') -Raw | ConvertFrom-Json
    $config = $project.batchRunner
    $dateRoot = Resolve-BatchPath $RepoRoot $config.dateRoot
    $cataloguePath = Resolve-BatchPath $RepoRoot $config.catalogue
    if ($SequenceProfile -ne 'Current') {
        if (-not $config.PSObject.Properties['sequenceProfiles']) { throw 'Sequence profiles are not configured.' }
        $cataloguePath = Resolve-BatchPath $RepoRoot $config.sequenceProfiles.$SequenceProfile
    }
    # Keep the historical project key/file path, now used only for batch settings.
    $settingsPath = Resolve-BatchPath $RepoRoot $config.approval
    $profilePath = Join-Path $dateRoot 'runner/ResidentialCare-ClientRunProfile.psd1'
    $unitPath = Join-Path $dateRoot 'runner/ResidentialCare-UnitWorkbookSequence.psd1'
    $orgPath = Join-Path $dateRoot 'runner/ResidentialCare-OrgWorkbookSequence.psd1'
    $catalogue = Import-PowerShellDataFile -LiteralPath $cataloguePath
    $settings = Import-PowerShellDataFile -LiteralPath $settingsPath
    $profile = Import-PowerShellDataFile -LiteralPath $profilePath
    $unitManifest = Import-PowerShellDataFile -LiteralPath $unitPath
    $orgManifest = Import-PowerShellDataFile -LiteralPath $orgPath
    if ($catalogue.SchemaVersion -ne 1 -or $profile.SchemaVersion -ne 1) { throw 'Unsupported catalogue or role profile version.' }
    if (@($profile.RoleWorkbookOrder).Count -ne 3) { throw 'Each role batch requires exactly three ordered workbooks.' }
    foreach ($file in $profile.RoleWorkbookOrder) {
        if ($file -match '[\\/:]' -or [IO.Path]::GetExtension($file) -ne '.xlsx') { throw "Invalid role workbook filename: $file" }
    }
    # JSON scope is opt-in until the unit and organisation workbooks are migrated.
    # With the project key absent, the existing production runner remains unchanged.
    $scopePath = $null
    $unitRoles = @{}
    $additionalUnitNames = @()
    if ($config.PSObject.Properties['runScope']) {
        $scopePath = Resolve-BatchPath $RepoRoot $config.runScope
        $scope = Read-ResidentialCareRunScope $scopePath 'DATExx-Whiddon'
        $unitNames = @($scope.Units)
        $roles = @($scope.Roles)
        $orgUnits = @($scope.Units)
        foreach ($unit in $unitNames) { $unitRoles[$unit] = @($roles) }
        Assert-ResidentialCareRunScopeFolders $scope (Join-Path $dateRoot 'UNITS') @($profile.RoleWorkbookOrder)
    } else {
        $legacyUnitNames = @(Get-ChildItem -LiteralPath (Join-Path $dateRoot 'UNITS') -Directory |
            Where-Object Name -match '^Unit[0-9]+$' | Sort-Object { [int] ($_.Name -replace '^Unit', '') } | ForEach-Object Name)
        if ($legacyUnitNames.Count -eq 0) { throw 'No Unit folders found.' }
        $legacyRoles = @($profile.Roles | Where-Object Enabled | ForEach-Object Folder)
        foreach ($entry in $profile.Roles) {
            if ($entry.Enabled -isnot [bool] -or $entry.Folder -isnot [string] -or [string]::IsNullOrWhiteSpace($entry.Folder)) { throw 'Each role requires a folder name and a boolean Enabled value.' }
        }
        if ($legacyRoles.Count -eq 0 -or @($legacyRoles | Select-Object -Unique).Count -ne $legacyRoles.Count) { throw 'Enabled roles must be nonempty and unique.' }
        foreach ($role in $legacyRoles) {
            if ($role -match '[\\/:]' -or $role -in @('.', '..')) { throw "Invalid role folder: $role" }
        }
        foreach ($unit in $legacyUnitNames) { $unitRoles[$unit] = @($legacyRoles) }
        $configuredAdditionalUnits = if ($settings.ContainsKey('AdditionalUnits')) { @($settings.AdditionalUnits) } else { @() }
        foreach ($entry in $configuredAdditionalUnits) {
            if ($entry -isnot [hashtable] -or $entry.Folder -isnot [string] -or
                $entry.Folder -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]*$') {
                throw 'AdditionalUnits must contain safe folder names and ordered roles.'
            }
            $unit = $entry.Folder
            if ($unit -in $legacyUnitNames -or $unit -in $additionalUnitNames) { throw "Duplicate additional Unit: $unit" }
            $candidateUnitPath = Join-Path $dateRoot "UNITS/$unit"
            if (-not (Test-Path -LiteralPath $candidateUnitPath -PathType Container)) { throw "Additional Unit folder is missing: $unit" }
            if ((Get-Item -LiteralPath $candidateUnitPath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Additional Unit folder is a reparse point: $unit"
            }
            $unitRoleNames = @($entry.Roles)
            if ($unitRoleNames.Count -eq 0 -or @($unitRoleNames | Select-Object -Unique).Count -ne $unitRoleNames.Count) {
                throw "Additional Unit roles must be nonempty and unique: $unit"
            }
            foreach ($role in $unitRoleNames) {
                if ($role -isnot [string] -or $role -cnotmatch '^[A-Za-z0-9][A-Za-z0-9_-]*$') {
                    throw "Invalid role folder for additional Unit $unit`: $role"
                }
            }
            $additionalUnitNames += $unit
            $unitRoles[$unit] = $unitRoleNames
        }
        $unitNames = @($legacyUnitNames) + @($additionalUnitNames)
        $roles = @($legacyRoles + @($additionalUnitNames | ForEach-Object { $unitRoles[$_] }) | Select-Object -Unique)
        $orgUnits = @($settings.OrganisationUnits)
        if ($orgUnits.Count -eq 0) { throw 'Declare organisation consumer Units in the batch settings file.' }
        foreach ($unit in $orgUnits) { if ($unit -notin $unitNames) { throw "Organisation consumer Unit not found: $unit" } }
    }
    $runAllUnits = if ($settings.ContainsKey('RunAllUnits')) { @($settings.RunAllUnits) } else { @($orgUnits) }
    if (-not $runAllUnits.Count -or @($runAllUnits | Select-Object -Unique).Count -ne $runAllUnits.Count) {
        throw 'RunAllUnits must be nonempty and unique.'
    }
    foreach ($unit in $runAllUnits) { if ($unit -notin $unitNames) { throw "RunAll Unit not found: $unit" } }
    $runAllIncludeOrg = if ($settings.ContainsKey('RunAllIncludeOrg')) { $settings.RunAllIncludeOrg } else { $true }
    if ($runAllIncludeOrg -isnot [bool]) { throw 'RunAllIncludeOrg must be boolean.' }
    $jobs = [Collections.Generic.List[object]]::new()
    $legacy = [Collections.Generic.List[object]]::new()
    $globalSequence = 0
    foreach ($unit in $unitNames) {
        $currentRoles = @($unitRoles[$unit])
        $legacyUnitEntries = if ($scopePath -or $unit -in $additionalUnitNames) {
            $entries = [Collections.Generic.List[object]]::new()
            foreach ($entry in @($unitManifest.Workbooks | Where-Object { $_.Sequence -le 13 } | Sort-Object Sequence)) {
                $entries.Add([pscustomobject]@{ Folder = $entry.Folder; FileName = $entry.FileName })
            }
            foreach ($role in $currentRoles) {
                foreach ($file in $profile.RoleWorkbookOrder) {
                    $entries.Add([pscustomobject]@{ Folder = "2. Calculations/$role"; FileName = $file })
                }
            }
            foreach ($entry in @($unitManifest.Workbooks | Where-Object { $_.Sequence -ge 23 } | Sort-Object Sequence)) {
                $entries.Add([pscustomobject]@{ Folder = $entry.Folder; FileName = $entry.FileName })
            }
            @($entries)
        } else { @($unitManifest.Workbooks | Sort-Object Sequence) }
        $localSequence = 0
        foreach ($entry in $legacyUnitEntries) {
            $globalSequence++
            $localSequence++
            $relative = "UNITS/$unit/$($entry.Folder)/$($entry.FileName)" -replace '\\', '/'
            $legacy.Add([pscustomobject]@{ Global = $globalSequence; Local = $localSequence; Unit = $unit; Path = $relative })
        }
        foreach ($spec in $catalogue.UnitJobs) {
            $localPath = if ($spec.ContainsKey('Sequence')) {
                $entry = @($unitManifest.Workbooks | Where-Object Sequence -eq $spec.Sequence)
                if ($entry.Count -ne 1) { throw "Invalid Unit manifest reference: $($spec.Sequence)" }
                "$($entry[0].Folder)/$($entry[0].FileName)"
            } else { $spec.Path }
            $dependencies = @($spec.Depends | ForEach-Object {
                if ($_ -eq 'RoleOutputs') { foreach ($role in $currentRoles) { "$unit/Role:$role/3" } }
                else { "$unit/$_" }
            })
            $inputs = if ($spec.ContainsKey('Inputs')) { @($spec.Inputs | ForEach-Object { "UNITS/$unit/$_" }) } else { @() }
            $jobs.Add([pscustomobject]@{
                Id = "$unit/$($spec.Id)"; Unit = $unit; Batch = $spec.Batch; BatchKey = "$unit/$($spec.Batch)"; Role = ''
                RelativePath = ("UNITS/$unit/$localPath" -replace '\\', '/'); Dependencies = $dependencies; InputPaths = $inputs
                BatchOrder = [array]::IndexOf($catalogue.BatchOrder, $spec.Batch); RoleOrder = 0; FileOrder = $jobs.Count
            })
        }
        for ($r = 0; $r -lt $currentRoles.Count; $r++) {
            $role = $currentRoles[$r]
            for ($f = 0; $f -lt 3; $f++) {
                $deps = @($catalogue.RoleDependencies | ForEach-Object { "$unit/$_" })
                if ($f -gt 0) { $deps += "$unit/Role:$role/$f" }
                $jobs.Add([pscustomobject]@{
                    Id = "$unit/Role:$role/$($f + 1)"; Unit = $unit; Batch = "C2.$($r + 1)"; BatchKey = "$unit/C2:$role"; Role = $role
                    RelativePath = "UNITS/$unit/2. Calculations/$role/$($profile.RoleWorkbookOrder[$f])"
                    Dependencies = $deps; InputPaths = @(); BatchOrder = [array]::IndexOf($catalogue.BatchOrder, 'C2'); RoleOrder = $r; FileOrder = $f
                })
            }
        }
    }
    foreach ($entry in ($orgManifest.Workbooks | Sort-Object Sequence)) {
        $globalSequence++
        $legacy.Add([pscustomobject]@{ Global = $globalSequence; Local = [int] $entry.Sequence; Unit = 'Org'; Path = ($entry.Path -replace '\\', '/') })
    }
    foreach ($spec in $catalogue.OrgJobs) {
        $entry = @($orgManifest.Workbooks | Where-Object Sequence -eq $spec.Sequence)
        if ($entry.Count -ne 1) { throw "Invalid organisation manifest reference: $($spec.Sequence)" }
        $deps = @($spec.Depends | ForEach-Object { "Org/$_" })
        foreach ($unit in $orgUnits) { $deps += @($spec.UnitDepends | ForEach-Object { "$unit/$_" }) }
        $jobs.Add([pscustomobject]@{
            Id = "Org/$($spec.Id)"; Unit = 'Org'; Batch = $spec.Batch; BatchKey = "Org/$($spec.Batch)"; Role = ''
            RelativePath = ($entry[0].Path -replace '\\', '/'); Dependencies = $deps
            InputPaths = $(if ($spec.ContainsKey('Inputs')) { @($spec.Inputs) } else { @() })
            BatchOrder = [array]::IndexOf($catalogue.BatchOrder, $spec.Batch); RoleOrder = 0; FileOrder = [int] $spec.Sequence
        })
    }
    $byId = @{}; $byPath = @{}
    foreach ($job in $jobs) {
        if ($byId.ContainsKey($job.Id) -or $byPath.ContainsKey($job.RelativePath)) { throw "Duplicate job or target: $($job.Id)" }
        $byId[$job.Id] = $job; $byPath[$job.RelativePath] = $job
        $job | Add-Member Path (Resolve-BatchPath $dateRoot $job.RelativePath)
    }
    if ($scopePath) {
        # An old organisation input must never read a disabled Unit merely
        # because that workbook still exists on disk.
        foreach ($job in $jobs) {
            foreach ($inputPath in @($job.InputPaths)) {
                if ($inputPath -match '^UNITS/([^/]+)/' -and $Matches[1] -notin $unitNames) {
                    throw "JSON scope cannot activate: $($job.Id) reads disabled Unit $($Matches[1]) ($inputPath)."
                }
            }
        }
    }
    foreach ($id in @($settings.AdditionalInputs.Keys) + @($settings.ExclusiveJobs)) {
        if (-not $byId.ContainsKey($id)) { throw "Batch settings reference an unknown job: $id" }
    }
    foreach ($job in $jobs) {
        # Input paths which are also producers must participate in scheduling.
        foreach ($path in @($job.InputPaths)) {
            if ($byPath.ContainsKey($path)) { $job.Dependencies += $byPath[$path].Id }
        }
        $job.Dependencies = @($job.Dependencies | Select-Object -Unique)
        $reads = @($job.InputPaths | ForEach-Object { Resolve-BatchPath $dateRoot $_ })
        # ReadJobs is independent of ordering edges. Existing catalogues retain
        # their implicit producer reads; explicit reads survive an ordering edit.
        $readJobs = @($job.Dependencies)
        if ($job.Unit -ne 'Org' -and -not $job.Role) {
            $definition = @($catalogue.UnitJobs | Where-Object { "$($job.Unit)/$($_.Id)" -eq $job.Id })[0]
            if ($definition.ContainsKey('ReadJobs')) {
                $readJobs = @($definition.ReadJobs | ForEach-Object {
                    if ($_ -eq 'RoleOutputs') { foreach ($role in @($unitRoles[$job.Unit])) { "$($job.Unit)/Role:$role/3" } }
                    else { "$($job.Unit)/$_" }
                })
            }
        }
        foreach ($dep in $readJobs) {
            if (-not $byId.ContainsKey($dep)) { throw "Unknown dependency $dep for $($job.Id)" }
            $reads += $byId[$dep].Path
        }
        if ($settings.AdditionalInputs.ContainsKey($job.Id)) {
            foreach ($inputPath in $settings.AdditionalInputs[$job.Id]) {
                if ($inputPath -is [hashtable]) {
                    $inputRoot = [Environment]::GetEnvironmentVariable($inputPath.Environment)
                    if (-not $inputRoot) { throw "Missing input-root environment variable: $($inputPath.Environment)" }
                    $reads += Resolve-BatchPath $inputRoot $inputPath.Path
                } else { $reads += Resolve-BatchPath $RepoRoot $inputPath }
            }
        }
        # Additional approved inputs may themselves refer to scheduled producers.
        foreach ($producer in $jobs) {
            if ($producer.Path -in $reads -and $producer.Id -ne $job.Id) { $job.Dependencies += $producer.Id }
        }
        $job.Dependencies = @($job.Dependencies | Select-Object -Unique)
        $job | Add-Member Reads @($reads | Where-Object { $_ -ne $job.Path } | Select-Object -Unique)
        $job | Add-Member Exclusive ([bool] ($job.Id -in $settings.ExclusiveJobs))
    }
    # Preserve batch sequencing as explicit edges, even for disjoint selections.
    foreach ($group in ($jobs | Group-Object BatchKey)) {
        $ordered = @($group.Group | Sort-Object FileOrder)
        for ($i = 1; $i -lt $ordered.Count; $i++) {
            $ordered[$i].Dependencies = @(($ordered[$i].Dependencies + $ordered[$i - 1].Id) | Select-Object -Unique)
        }
    }
    $seen = @{}
    while ($seen.Count -lt $jobs.Count) {
        $ready = @($jobs | Where-Object { -not $seen.ContainsKey($_.Id) -and @($_.Dependencies | Where-Object { -not $seen.ContainsKey($_) }).Count -eq 0 })
        if ($ready.Count -eq 0) { throw 'Batch catalogue contains a dependency cycle.' }
        foreach ($job in $ready) { $seen[$job.Id] = $true }
    }
    $files = @((Join-Path $RepoRoot 'pq.project.json'), $cataloguePath, $settingsPath, $profilePath, $unitPath, $orgPath)
    if ($scopePath) { $files += $scopePath }
    $files += @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter 'Batch*.ps1' -File | Sort-Object Name | ForEach-Object FullName)
    $files += @('RunScope.ps1', 'BatchScheduling.ps1', 'RefreshJson.ps1', 'RefreshGitAccess.ps1', 'Invoke-BatchWorkbook.ps1', 'RefreshRunGate.ps1', 'ExcelRefreshSafety.ps1', 'Invoke-AllUnitsExcelWorkbookRefresh.ps1' | ForEach-Object { Join-Path $PSScriptRoot $_ })
    $fingerprint = Get-BatchFingerprint $files
    return [pscustomobject]@{
        SchemaVersion = 1; DateRoot = $dateRoot; LogRoot = (Resolve-BatchPath $RepoRoot $config.logFolder)
        SequenceProfile = $SequenceProfile
        Units = $unitNames; Roles = $roles; OrgUnits = $orgUnits
        RunAllUnits = $runAllUnits; RunAllIncludeOrg = $runAllIncludeOrg; Settings = $settings; Fingerprint = $fingerprint
        BatchTitles = $(if ($catalogue.ContainsKey('BatchTitles')) { $catalogue.BatchTitles } else { @{} })
        Jobs = @($jobs | Sort-Object @{Expression={ if ($_.Unit -eq 'Org') { [int]::MaxValue } else { [array]::IndexOf($unitNames, $_.Unit) } }}, BatchOrder, RoleOrder, FileOrder)
        Legacy = @($legacy); WorkerScript = (Join-Path $dateRoot 'runner/src/Invoke-AllUnitsExcelWorkbookRefresh.ps1')
    }
}

function Get-UnitPickList {
    param($Catalogue, [switch] $BlankMeansAll)
    ''
    'Available Unit IDs:'
    foreach ($unit in $Catalogue.Units) { $unit }
    'All - all listed Units'
    ''
    $example = @($Catalogue.Units | Select-Object -First 2) -join ','
    "Pick one or more Unit IDs, comma-separated (for example: $example), or All."
    if ($BlankMeansAll) { 'Blank = all listed Units.' }
    else { 'Blank = cancel this selection; nothing will run.' }
}

function Read-UnitSelection {
    param($Catalogue, [switch] $BlankMeansAll)
    foreach ($line in (Get-UnitPickList $Catalogue -BlankMeansAll:$BlankMeansAll)) { Write-Host $line }
    $prompt = if ($BlankMeansAll) { 'Unit IDs, comma-separated (or All; blank = all)' }
        else { 'Unit IDs, comma-separated (or All; blank = cancel)' }
    Read-Host $prompt
}

function Get-BatchPickList {
    param($Catalogue)
    # Batch IDs apply across Units: show each choice once, using its actual
    # manifest filenames in execution order rather than repeating Unit folders.
    $unitJobs = @($Catalogue.Jobs | Where-Object Unit -eq $Catalogue.Units[0])
    $orgJobs = @($Catalogue.Jobs | Where-Object Unit -eq 'Org')
    foreach ($section in @(
        @{ Heading = 'Unit batches (applied to the Units you select next):'; Jobs = $unitJobs },
        @{ Heading = 'Organisation batches (run once):'; Jobs = $orgJobs }
    )) {
        ''
        $section.Heading
        $shownRoleAlias = $false
        foreach ($id in @($section.Jobs | ForEach-Object Batch | Select-Object -Unique)) {
            $batchJobs = @($section.Jobs | Where-Object Batch -eq $id)
            $files = @($batchJobs | ForEach-Object { Split-Path $_.RelativePath -Leaf }) -join ', '
            $role = $batchJobs[0].Role
            $titleKey = if ($role) { 'C2' } else { $id }
            $title = if ($Catalogue.BatchTitles.ContainsKey($titleKey)) { [string] $Catalogue.BatchTitles[$titleKey] } else { '' }
            if ($role) {
                if (-not $shownRoleAlias) {
                    "C2 - All enabled role capacity batches [$files]"
                    $shownRoleAlias = $true
                }
                $roleMappings = @($Catalogue.Jobs | Where-Object { $_.Unit -ne 'Org' -and $_.Batch -eq $id } |
                    Group-Object Role | ForEach-Object {
                        "$($_.Name) ($(@($_.Group | Select-Object -ExpandProperty Unit -Unique) -join ', '))"
                    })
                $roleLabel = if ($roleMappings.Count -eq 1) { $role } else { $roleMappings -join ' / ' }
                $title = if ($title) { "$title - $roleLabel" } else { $roleLabel }
            }
            $label = if ($title) { "$id - $title" } else { $id }
            "$label [$files]"
        }
    }
    ''
    'Pick one or more batch IDs, comma-separated (for example: D1,A2,C2.1). C2 selects every enabled role batch.'
    'Organisation IDs start with the letter O (for example O5); 05 with a zero is also accepted.'
}

function Test-BatchSelectionNeedsUnits {
    param($Catalogue, [string[]] $Batches)
    foreach ($batch in (Expand-BatchArguments $Batches)) {
        if ($batch -eq 'C2' -or @($Catalogue.Jobs | Where-Object { $_.Unit -ne 'Org' -and $_.Batch -eq $batch }).Count) {
            return $true
        }
    }
    return $false
}

function Select-BatchPlan {
    param($Catalogue, [switch] $RunAll, [string[]] $Units, [string[]] $Batches, [string[]] $Roles,
        [string[]] $Workbooks, [string] $StartAtWorkbook, [int] $StartAtSequence, [int] $EndAtSequence,
        [switch] $IncludeOrg, [switch] $IncludeDependencies)
    $Units = @(Expand-BatchArguments $Units); $Batches = @(Expand-BatchArguments $Batches)
    $Batches = @($Batches | ForEach-Object {
        if ($_ -match '^0([1-5])$') {
            $orgBatch = "O$($Matches[1])"
            Write-Host "Interpreting batch ID $_ as $orgBatch (letter O)."
            $orgBatch
        } else { $_ }
    } | Select-Object -Unique)
    $Roles = @(Expand-BatchArguments $Roles); $Workbooks = @(Expand-BatchArguments $Workbooks)
    $modes = [int] [bool] $RunAll + [int] [bool] $Batches.Count + [int] [bool] $Workbooks.Count + [int] [bool] $StartAtWorkbook + [int] [bool] ($StartAtSequence -or $EndAtSequence)
    if ($modes -gt 1) { throw 'Choose one selection mode: all, batches, workbooks, start-at-workbook, or sequence range.' }
    if ($Roles.Count -and ($RunAll -or $Workbooks.Count -or $StartAtWorkbook -or $StartAtSequence -or $EndAtSequence)) { throw 'Roles may qualify C2 batches or be selected on their own.' }
    if ($Roles.Count -and $Batches.Count -and -not @($Batches | Where-Object { $_ -eq 'C2' -or $_ -match '^C2\.[0-9]+$' }).Count) { throw 'A role filter requires a C2 batch selection.' }
    $explicitUnits = $Units.Count -gt 0
    if ($Units -contains 'All') {
        if ($Units.Count -ne 1) { throw 'Use All alone in the Unit selector.' }
        $Units = @($Catalogue.Units)
    }
    # Unqualified run and preview actions follow the configured default scope.
    $defaultRunScope = -not $explicitUnits -and ($RunAll -or -not ($Batches.Count -or $Roles.Count -or
        $Workbooks.Count -or $StartAtWorkbook -or $StartAtSequence -or $EndAtSequence))
    if (-not $Units.Count) {
        $Units = if ($defaultRunScope) { @($Catalogue.RunAllUnits) } else { @($Catalogue.Units) }
    }
    foreach ($unit in $Units) { if ($unit -notin $Catalogue.Units) { throw "Unknown Unit: $unit" } }
    foreach ($role in $Roles) { if ($role -notin $Catalogue.Roles) { throw "Unknown or disabled role: $role" } }
    $eligible = @($Catalogue.Jobs | Where-Object { $_.Unit -in $Units -or $_.Unit -eq 'Org' })
    $selected = @()
    if ($StartAtSequence -or $EndAtSequence -or $StartAtWorkbook) {
        $legacy = @($Catalogue.Legacy)
        $number = 'Global'
        if ($explicitUnits) {
            if ($Units.Count -ne 1) { throw 'A Unit-local legacy range requires exactly one Unit.' }
            $legacy = @($legacy | Where-Object Unit -eq $Units[0]); $number = 'Local'
        }
        if ($StartAtWorkbook) {
            $match = @($legacy | Where-Object { $_.Path -eq ($StartAtWorkbook -replace '\\', '/') -or (Split-Path $_.Path -Leaf) -eq $StartAtWorkbook })
            if ($match.Count -ne 1) { throw 'Start workbook is missing or ambiguous; use a Date-relative path and a legacy workbook.' }
            $StartAtSequence = $match[0].$number
        }
        if ($StartAtSequence -lt 1) { throw 'StartAtSequence must be positive.' }
        if (-not $EndAtSequence) { $EndAtSequence = $legacy[-1].$number }
        if ($EndAtSequence -lt $StartAtSequence -or $EndAtSequence -gt $legacy[-1].$number) { throw 'Invalid legacy range.' }
        foreach ($entry in ($legacy | Where-Object { $_.$number -ge $StartAtSequence -and $_.$number -le $EndAtSequence })) {
            $match = @($eligible | Where-Object RelativePath -eq $entry.Path)
            if ($match.Count -ne 1) { throw "Legacy target is no longer enabled: $($entry.Path)" }
            $selected += $match[0]
        }
    } elseif ($Workbooks.Count) {
        foreach ($file in $Workbooks) {
            $match = @($eligible | Where-Object { $_.RelativePath -eq ($file -replace '\\', '/') -or (Split-Path $_.RelativePath -Leaf) -eq $file -or $_.Id -eq $file })
            if ($match.Count -ne 1) { throw "Workbook selection is missing or ambiguous: $file. Use its Date-relative path." }
            $selected += $match[0]
        }
    } elseif ($Batches.Count -or $Roles.Count) {
        if (-not $Batches.Count) { $Batches = @('C2') }
        foreach ($batch in $Batches) {
            $match = @($eligible | Where-Object { ($_.Batch -eq $batch -or ($batch -eq 'C2' -and $_.Role)) -and (-not $_.Role -or -not $Roles.Count -or $_.Role -in $Roles) })
            if (-not $match.Count) { throw "Unknown or empty batch selection: $batch" }
            $selected += $match
        }
    } else {
        $selected = @($eligible | Where-Object {
            $_.Unit -ne 'Org' -or ($defaultRunScope -and $Catalogue.RunAllIncludeOrg) -or $IncludeOrg
        })
    }
    if ($IncludeOrg) { $selected += @($Catalogue.Jobs | Where-Object Unit -eq 'Org') }
    $ids = @{}; foreach ($job in $selected) { $ids[$job.Id] = $true }
    if ($IncludeDependencies) {
        do {
            $previousCount = $ids.Count
            foreach ($job in $Catalogue.Jobs) {
                if ($ids.ContainsKey($job.Id)) { foreach ($dep in $job.Dependencies) { $ids[$dep] = $true } }
            }
        } while ($ids.Count -gt $previousCount)
    }
    $selected = @($Catalogue.Jobs | Where-Object { $ids.ContainsKey($_.Id) })
    if (-not $selected.Count) { throw 'The selection is empty.' }
    $unitOnlyJobs = @($selected | Where-Object { $_.Unit -ne 'Org' -and $_.Unit -notin $Catalogue.OrgUnits })
    if ($unitOnlyJobs.Count -and @($selected | Where-Object Unit -eq 'Org').Count) {
        $unitOnlyNames = @($unitOnlyJobs | Select-Object -ExpandProperty Unit -Unique) -join ', '
        throw "Organisation jobs do not consume $unitOnlyNames yet. Select those Units without organisation batches."
    }
    return [pscustomobject]@{
        SchemaVersion = 1; Fingerprint = $Catalogue.Fingerprint; DateRoot = $Catalogue.DateRoot
        SequenceProfile = $Catalogue.SequenceProfile
        Roles = $Catalogue.Roles; OrgUnits = $Catalogue.OrgUnits; Jobs = $selected
    }
}

function Get-BatchFileStamp {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return 'missing' }
    $item = Get-Item -LiteralPath $Path
    return "$($item.Length):$($item.LastWriteTimeUtc.Ticks)"
}

function Assert-BatchFileAccessible {
    param([string] $Path, [switch] $OutputFile, [switch] $InspectFileUsers,
        [ValidateRange(0, 60)] [double] $FileUserWaitSeconds = 0)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing file: $Path" }
    $lock = Join-Path (Split-Path $Path -Parent) ('~$' + (Split-Path $Path -Leaf))
    if (Test-Path -LiteralPath $lock) { throw "Excel lock file found: $lock" }
    if ($InspectFileUsers -and $OutputFile) {
        Assert-NoExternalFileUsers -Path $Path -WaitSeconds $FileUserWaitSeconds -OnWait {
            param($description)
            Write-Host "Waiting for workbook access before dispatch: $description"
        }
    }
    $access = if ($OutputFile) { [IO.FileAccess]::ReadWrite } else { [IO.FileAccess]::Read }
    $stream = [IO.File]::Open($Path, 'Open', $access, 'Read')
    $stream.Dispose()
}

function Test-BatchPlan {
    param($Plan, [switch] $InspectFileUsers)
    $issues = [Collections.Generic.List[string]]::new()
    $checked = @{}
    foreach ($job in $Plan.Jobs) {
        try { Assert-BatchFileAccessible $job.Path -OutputFile -InspectFileUsers:$InspectFileUsers }
        catch { $issues.Add("$($job.Id): $($_.Exception.Message)") }
        foreach ($path in $job.Reads) {
            if ($checked.ContainsKey($path)) { continue }; $checked[$path] = $true
            try { Assert-BatchFileAccessible $path }
            catch { $issues.Add("$($job.Id) input: $($_.Exception.Message)") }
        }
    }
    return @($issues)
}
