[CmdletBinding()]
param(
    [string] $SourceRoot = 'Workflows/ResidentialCare/CapacityDistribution',
    [switch] $RunNative
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$resolvedSourceRoot = Join-Path $repoRoot $SourceRoot
$sourcePaths = [ordered]@{
    A1 = Join-Path $resolvedSourceRoot 'CapacityDistrib-A1.pq'
    A2 = Join-Path $resolvedSourceRoot 'CapacityDistrib-A2.pq'
    B  = Join-Path $resolvedSourceRoot 'CapacityDistrib-B.pq'
}
$results = [Collections.Generic.List[object]]::new()

function Add-ContractResult {
    param(
        [Parameter(Mandatory)][string] $Name,
        [Parameter(Mandatory)][string] $File,
        [Parameter(Mandatory)][bool] $Passed,
        [Parameter(Mandatory)][string] $Detail
    )
    $results.Add([pscustomobject]@{ Check = $Name; File = $File; Passed = $Passed; Detail = $Detail })
}

function Get-MSharedDefinition {
    param(
        [Parameter(Mandatory)][string] $Text,
        [Parameter(Mandatory)][string] $Name
    )

    $escapedQuotedName = [regex]::Escape($Name.Replace('"', '""'))
    $quotedPattern = '(?m)^shared\s+#"' + $escapedQuotedName + '"\s*='
    $plainPattern = '(?m)^shared\s+' + [regex]::Escape($Name) + '\s*='
    $match = [regex]::Match($Text, $quotedPattern)
    if (-not $match.Success) { $match = [regex]::Match($Text, $plainPattern) }
    if (-not $match.Success) { return $null }

    $inString = $false
    $lineComment = $false
    $blockComment = $false
    for ($index = $match.Index; $index -lt $Text.Length; $index++) {
        $current = $Text[$index]
        $next = if ($index + 1 -lt $Text.Length) { $Text[$index + 1] } else { [char]0 }

        if ($lineComment) {
            if ($current -eq "`n") { $lineComment = $false }
            continue
        }
        if ($blockComment) {
            if ($current -eq '*' -and $next -eq '/') { $blockComment = $false; $index++ }
            continue
        }
        if ($inString) {
            if ($current -eq '"' -and $next -eq '"') { $index++ }
            elseif ($current -eq '"') { $inString = $false }
            continue
        }
        if ($current -eq '/' -and $next -eq '/') { $lineComment = $true; $index++; continue }
        if ($current -eq '/' -and $next -eq '*') { $blockComment = $true; $index++; continue }
        if ($current -eq '"') { $inString = $true; continue }
        if ($current -eq ';') { return $Text.Substring($match.Index, $index - $match.Index + 1) }
    }

    return $null
}

function Get-MDeclarationNames {
    param([Parameter(Mandatory)][string] $Text)
    $matches = [regex]::Matches($Text, '(?m)^shared\s+(?<name>#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=')
    foreach ($match in $matches) {
        $rawName = $match.Groups['name'].Value
        if ($rawName.StartsWith('#"')) { $rawName.Substring(2, $rawName.Length - 3).Replace('""', '"') } else { $rawName }
    }
}

function Get-CompactM {
    param([AllowNull()][string] $Text)
    if ($null -eq $Text) { return '' }
    return [regex]::Replace($Text, '\s+', ' ').Trim()
}

function Test-MFinalSortKey {
    param(
        [AllowNull()][string] $Definition,
        [Parameter(Mandatory)][string] $ExpectedKey
    )
    if ([string]::IsNullOrWhiteSpace($Definition)) { return $false }
    $sortMatches = [regex]::Matches($Definition, 'Table\.Sort\([\s\S]*?,\s*\{\{(?<spec>[\s\S]*?)\}\}\s*\)')
    foreach ($sortMatch in $sortMatches) {
        $keyMatches = [regex]::Matches($sortMatch.Groups['spec'].Value, '"(?<key>(?:[^"]|"")+)"\s*,\s*Order\.(?:Ascending|Descending)')
        if ($keyMatches.Count -gt 0 -and $keyMatches[$keyMatches.Count - 1].Groups['key'].Value.Replace('""', '"') -eq $ExpectedKey) {
            return $true
        }
    }
    return $false
}

function Test-MSortContainsKey {
    param(
        [AllowNull()][string] $Definition,
        [Parameter(Mandatory)][string] $ExpectedKey
    )
    if ([string]::IsNullOrWhiteSpace($Definition)) { return $false }
    $sortMatches = [regex]::Matches($Definition, 'Table\.Sort\([\s\S]*?,\s*\{\{(?<spec>[\s\S]*?)\}\}\s*\)')
    foreach ($sortMatch in $sortMatches) {
        $keyMatches = [regex]::Matches($sortMatch.Groups['spec'].Value, '"(?<key>(?:[^"]|"")+)"\s*,\s*Order\.(?:Ascending|Descending)')
        if (@($keyMatches | Where-Object { $_.Groups['key'].Value.Replace('""', '"') -eq $ExpectedKey }).Count -gt 0) {
            return $true
        }
    }
    return $false
}

function Add-DefinitionPresence {
    param([string] $Stage, [string] $Text, [string] $Name)
    $definition = Get-MSharedDefinition -Text $Text -Name $Name
    Add-ContractResult -Name "shared:$Name" -File $Stage -Passed ($null -ne $definition) -Detail $(
        if ($null -ne $definition) { 'Required shared definition is present.' } else { 'Required shared definition is missing.' }
    )
    return $definition
}

$sources = @{}
foreach ($entry in $sourcePaths.GetEnumerator()) {
    $exists = Test-Path -LiteralPath $entry.Value -PathType Leaf
    Add-ContractResult -Name 'source-file' -File $entry.Key -Passed $exists -Detail $(
        if ($exists) { "Found $($entry.Value.Substring($repoRoot.Length + 1))." } else { "Missing $($entry.Value.Substring($repoRoot.Length + 1))." }
    )
    if ($exists) {
        $sources[$entry.Key] = [IO.File]::ReadAllText($entry.Value).Replace("`r`n", "`n")
    }
}

if ($sources.Count -eq $sourcePaths.Count) {
    foreach ($stage in @('A1', 'A2', 'B')) {
        $text = $sources[$stage]
        $names = @(Get-MDeclarationNames -Text $text)
        $duplicates = @($names | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name)
        Add-ContractResult -Name 'unique-shared-definitions' -File $stage -Passed ($duplicates.Count -eq 0) -Detail $(
            if ($duplicates.Count -eq 0) { "All $($names.Count) shared names are unique." } else { 'Duplicate definitions: ' + ($duplicates -join ', ') }
        )

        $roleFolder = Add-DefinitionPresence -Stage $stage -Text $text -Name 'RoleFolder'
        $sourceRole = Add-DefinitionPresence -Stage $stage -Text $text -Name 'SourceRole'
        $outputRole = Add-DefinitionPresence -Stage $stage -Text $text -Name 'OutputRole'
        $roleAlias = Add-DefinitionPresence -Stage $stage -Text $text -Name 'Role'
        $roleCheck = Add-DefinitionPresence -Stage $stage -Text $text -Name 'RoleContext_CHECK'
        $rolePath = Add-DefinitionPresence -Stage $stage -Text $text -Name 'RolePath'

        if ($null -ne $sourceRole) {
            $compact = Get-CompactM $sourceRole
            Add-ContractResult -Name 'source-role-mapping' -File $stage -Passed (
                $compact.Contains('RoleFolder') -and $compact.Contains('"AIN2"') -and $compact.Contains('"AIN"')
            ) -Detail 'SourceRole must derive AIN from the physical AIN2 or AIN role folder.'
        }
        if ($null -ne $outputRole) {
            $compact = Get-CompactM $outputRole
            Add-ContractResult -Name 'output-role-mapping' -File $stage -Passed (
                $compact.Contains('RoleFolder') -and $compact.Contains('"AIN2"') -and $compact.Contains('"AIN"')
            ) -Detail 'OutputRole must publish AIN2 during development and AIN after promotion.'
        }
        if ($null -ne $roleAlias) {
            Add-ContractResult -Name 'role-alias' -File $stage -Passed ((Get-CompactM $roleAlias).Contains('OutputRole')) -Detail 'Compatibility Role must resolve from OutputRole.'
        }
        if ($null -ne $roleCheck) {
            $compact = Get-CompactM $roleCheck
            Add-ContractResult -Name 'role-tuple-gate' -File $stage -Passed (
                $compact.Contains('RoleFolder') -and $compact.Contains('SourceRole') -and $compact.Contains('OutputRole') -and
                $compact.Contains('"AIN2"') -and $compact.Contains('"AIN"') -and
                ($compact.Contains('Error.Record') -or ($compact.Contains('Status') -and $compact.Contains('"Fail"')))
            ) -Detail 'RoleContext_CHECK must expose a failure for any tuple other than AIN2/AIN/AIN2 or AIN/AIN/AIN.'
        }
        if ($null -ne $rolePath) {
            $compact = Get-CompactM $rolePath
            Add-ContractResult -Name 'physical-role-path' -File $stage -Passed (
                $compact.Contains('RolePathTABLE') -and -not $compact.Contains('SourceRole') -and -not $compact.Contains('OutputRole')
            ) -Detail 'RolePath must remain the physical workbook folder and must not be redirected to the source role.'
        }
    }

    $a1 = $sources['A1']
    $a2 = $sources['A2']
    $b = $sources['B']

    foreach ($queryName in @(
        'IMPORT ShiftUnitDemandHRS !!',
        'IMPORT Masterlist !!',
        'IMPORT ResDayShift !!',
        'IMPORT ResourceShiftAllocation - Role!!',
        'EXTRACT PermutationDimensions'
    )) {
        $definition = Add-DefinitionPresence -Stage 'A1' -Text $a1 -Name $queryName
        if ($null -ne $definition) {
            $compact = Get-CompactM $definition
            Add-ContractResult -Name "source-output-role:$queryName" -File 'A1' -Passed (
                $compact.Contains('SourceRole') -and $compact.Contains('OutputRole')
            ) -Detail 'The import must filter upstream AIN rows with SourceRole and relabel working role values to OutputRole.'
        }
    }

    $contractImport = Add-DefinitionPresence -Stage 'A1' -Text $a1 -Name 'IMPORT Masterlist !!'
    $contractCheck = Add-DefinitionPresence -Stage 'A1' -Text $a1 -Name 'A1ResourceContract_CHECK'
    $contractCombined = (Get-CompactM $contractImport) + ' ' + (Get-CompactM $contractCheck)
    Add-ContractResult -Name 'a1-contract-role-semantics' -File 'A1' -Passed (
        $contractCombined.Contains('SourceRole') -and $contractCombined.Contains('OutputRole') -and $contractCombined.Contains('PreferredRole')
    ) -Detail 'Resource contracts must use OutputRole for published Role and SourceRole for PreferredRole selection.'

    $clusters = Add-DefinitionPresence -Stage 'A1' -Text $a1 -Name 'Clusters'
    if ($null -ne $clusters) {
        $compact = Get-CompactM $clusters
        $hasCompositeLeftKey = $compact.Contains('{"Resource", "PrevIndex"}') -or $compact.Contains('{"Resource","PrevIndex"}')
        $hasCompositeRightKey = $compact.Contains('{"Resource", "ResIndex"}') -or $compact.Contains('{"Resource","ResIndex"}') -or
            $compact.Contains('{"Resource", "LookupIndex"}') -or $compact.Contains('{"Resource","LookupIndex"}')
        $hasJoin = $compact.Contains('Table.NestedJoin(') -or $compact.Contains('Table.Join(')
        $hasLegacyScan = [regex]::IsMatch($compact, 'Table\.SelectRows\([^,]+,\s*each\s*\[ResIndex\]\s*=\s*PrevIndex')
        Add-ContractResult -Name 'resource-keyed-cluster-predecessor' -File 'A1' -Passed (
            $hasJoin -and $hasCompositeLeftKey -and $hasCompositeRightKey -and -not $hasLegacyScan
        ) -Detail 'Clusters must join Resource+PrevIndex to Resource+ResIndex and must not scan the complete table by index alone.'
    }

    $wdType = Add-DefinitionPresence -Stage 'A1' -Text $a1 -Name 'WD Type'
    if ($null -ne $wdType) {
        Add-ContractResult -Name 'nwd-fields-not-string-literals' -File 'A1' -Passed (
            $wdType.Contains('[Yesterday.Allocation3]') -and $wdType.Contains('[Tomorrow.Allocation3]') -and
            -not $wdType.Contains('"Yesterday.Allocation3"') -and -not $wdType.Contains('"Tomorrow.Allocation3"') -and
            -not $wdType.Contains('"Yesterday.Alloction3"') -and -not $wdType.Contains('"Tomorrow.Alloction3"')
        ) -Detail 'Conditional NWD rules must read the explicitly derived three-day allocation fields rather than compare quoted field-name text.'
    }
    $potential = Add-DefinitionPresence -Stage 'A1' -Text $a1 -Name 'ResDayPotentialAvailabilityTABLE'
    if ($null -ne $potential) {
        Add-ContractResult -Name 'nwd-exclusion-list' -File 'A1' -Passed ((Get-CompactM $potential).Contains('not List.Contains(')) -Detail 'NWD exclusions must use one explicit not-contains predicate rather than OR-combined not-equals tests.'
    }
    $multiPeriod = Add-DefinitionPresence -Stage 'A1' -Text $a1 -Name 'MultiPeriodPriority'
    if ($null -ne $multiPeriod) {
        Add-ContractResult -Name 'stable-sort:MultiPeriodPriority' -File 'A1' -Passed (Test-MSortContainsKey -Definition $multiPeriod -ExpectedKey 'AvailablePeriod') -Detail 'The priority sort must include AvailablePeriod as a deterministic key.'
    }

    $helper = Add-DefinitionPresence -Stage 'A2' -Text $a2 -Name 'fnAddIndexedRunningTotal'
    if ($null -ne $helper) {
        $compact = Get-CompactM $helper
        Add-ContractResult -Name 'a2-linear-helper' -File 'A2' -Passed (
            $compact.Contains('List.Generate(') -and $compact.Contains('List.Buffer(') -and
            -not $compact.Contains('List.Range(') -and -not $compact.Contains('Table.Sort(') -and -not $compact.Contains('Table.Group(')
        ) -Detail 'A2 must use a buffered-list, one-pass helper that does not reorder or regroup its input.'
    }

    $a2RunningQueries = @(
        @{ Name = 'Resource Running total'; Index = 'ResIndex'; Start = 'StartResIndex' },
        @{ Name = 'Period Running Total'; Index = 'PeriodIndex'; Start = 'StartPeriodIndex' }
    )
    foreach ($spec in $a2RunningQueries) {
        $definition = Add-DefinitionPresence -Stage 'A2' -Text $a2 -Name $spec.Name
        if ($null -ne $definition) {
            $compact = (Get-CompactM $definition).Replace(' ', '')
            $expectedCount = '[' + $spec.Index + ']-[' + $spec.Start + ']+1'
            Add-ContractResult -Name "linear-running-total:$($spec.Name)" -File 'A2' -Passed (
                $compact.Contains('fnAddIndexedRunningTotal(') -and -not $compact.Contains('List.Range(') -and $compact.Contains($expectedCount)
            ) -Detail 'Running totals must use the linear helper with count = current two-based index - two-based group start + 1.'
        }
    }

    foreach ($spec in @(
        @{ Name = 'ResCapIndexLimits'; Index = 'ResIndex'; Retired = 'ResIndexX' },
        @{ Name = 'PeriodCapIndexLimits'; Index = 'PeriodIndex'; Retired = 'PeriodIndexX' }
    )) {
        $definition = Add-DefinitionPresence -Stage 'A2' -Text $a2 -Name $spec.Name
        if ($null -ne $definition) {
            $compact = (Get-CompactM $definition).Replace(' ', '')
            Add-ContractResult -Name "two-based-group-start:$($spec.Name)" -File 'A2' -Passed (
                $compact.Contains(('List.Min([' + $spec.Index + '])')) -and -not $compact.Contains($spec.Retired)
            ) -Detail 'Group starts must come directly from the same two-based index consumed by the helper.'
        }
    }

    $capped = Add-DefinitionPresence -Stage 'A2' -Text $a2 -Name 'ResPeriodCappedAvailability(C#)TABLE'
    if ($null -ne $capped) {
        $compact = Get-CompactM $capped
        $skeletonAssignment = [regex]::IsMatch($compact, '(?:Source|Skeleton)\s*=\s*#"IMPORT ResourcePeriodTABLE_empty"')
        $expandedSkeletonRole = $compact.Contains('#"IMPORT ResourcePeriodTABLE_empty"') -and
            $compact.Contains('JoinKind.RightOuter') -and
            [regex]::IsMatch($compact, 'Table\.ExpandTableColumn\([^;]*"ResourcePeriodTABLE_empty"[^;]*"Role"')
        Add-ContractResult -Name 'a2-skeleton-owned-role' -File 'A2' -Passed (
            ($skeletonAssignment -or $expandedSkeletonRole) -and $compact.Contains('Role') -and -not $compact.Contains('Table.FillDown(')
        ) -Detail 'The complete skeleton must own Role; positional FillDown is forbidden.'
    }

    foreach ($spec in @(
        @{ Name = 'ResCapPrioritisedTABLE'; Key = 'Period' },
        @{ Name = 'PeriodCapPrioritisedTABLE'; Key = 'Resource' }
    )) {
        $definition = Add-DefinitionPresence -Stage 'A2' -Text $a2 -Name $spec.Name
        if ($null -ne $definition) {
            Add-ContractResult -Name "stable-sort:$($spec.Name)" -File 'A2' -Passed (Test-MSortContainsKey -Definition $definition -ExpectedKey $spec.Key) -Detail "The priority sort must include $($spec.Key) as a deterministic key."
        }
    }

    foreach ($name in @('ResPeriodAvailabilityCapped(C#)TABLE', 'C##TABLE', 'C###TABLE B')) {
        $null = Add-DefinitionPresence -Stage 'B' -Text $b -Name $name
    }
    foreach ($name in @('ResCapPrioritisedTABLE', 'Resource Running total', 'PeriodCapPrioritisedTABLE', 'Period Running Total', 'ResPeriodCappedAvailability(C#)TABLE')) {
        $null = Add-DefinitionPresence -Stage 'B' -Text $b -Name $name
    }

    foreach ($spec in @(
        @{ Name = 'Resource Running total'; Index = 'ResIndex'; Start = 'StartResIndex' },
        @{ Name = 'Period Running Total'; Index = 'PeriodIndex'; Start = 'StartPeriodIndex' }
    )) {
        $definition = Get-MSharedDefinition -Text $b -Name $spec.Name
        if ($null -ne $definition) {
            $compact = (Get-CompactM $definition).Replace(' ', '')
            $expectedCount = '[' + $spec.Index + ']-[' + $spec.Start + ']+1'
            Add-ContractResult -Name "merged-linear-running-total:$($spec.Name)" -File 'B' -Passed (
                $compact.Contains('fnAddIndexedRunningTotal(') -and -not $compact.Contains('List.Range(') -and $compact.Contains($expectedCount)
            ) -Detail 'Merged running totals must retain the tested linear helper and normalized two-based prefix count.'
        }
    }

    foreach ($spec in @(
        @{ Name = 'ResCapIndexLimits'; Index = 'ResIndex'; Retired = 'ResIndexX' },
        @{ Name = 'PeriodCapIndexLimits'; Index = 'PeriodIndex'; Retired = 'PeriodIndexX' }
    )) {
        $definition = Add-DefinitionPresence -Stage 'B' -Text $b -Name $spec.Name
        if ($null -ne $definition) {
            $compact = (Get-CompactM $definition).Replace(' ', '')
            Add-ContractResult -Name "merged-two-based-group-start:$($spec.Name)" -File 'B' -Passed (
                $compact.Contains(('List.Min([' + $spec.Index + '])')) -and -not $compact.Contains($spec.Retired)
            ) -Detail 'Merged group starts must come from the same two-based index consumed by the helper.'
        }
    }

    $mergedRawCapped = Get-MSharedDefinition -Text $b -Name 'ResPeriodCappedAvailability(C#)TABLE'
    if ($null -ne $mergedRawCapped) {
        $compact = Get-CompactM $mergedRawCapped
        $expandedSkeletonRole = $compact.Contains('#"IMPORT ResourcePeriodTABLE_empty"') -and
            $compact.Contains('JoinKind.RightOuter') -and
            [regex]::IsMatch($compact, 'Table\.ExpandTableColumn\([^;]*"ResourcePeriodTABLE_empty"[^;]*"Role"')
        Add-ContractResult -Name 'merged-b-skeleton-owned-role' -File 'B' -Passed (
            $expandedSkeletonRole -and -not $compact.Contains('Table.FillDown(')
        ) -Detail 'Merged C# must retain authoritative skeleton Role assignment and forbid positional FillDown.'
    }

    $retiredA2Import = Get-MSharedDefinition -Text $b -Name 'IMPORT ResPeriodAvailabilityCapped(C#)'
    Add-ContractResult -Name 'merged-b-no-a2-import-query' -File 'B' -Passed ($null -eq $retiredA2Import) -Detail 'Merged B must consume its internal C# stage rather than an A2 workbook query.'
    Add-ContractResult -Name 'merged-b-no-a2-workbook-path' -File 'B' -Passed (-not $b.Contains('CapacityDistrib(A.2)-shifts.xlsx')) -Detail 'Merged B must not open the retired A2 workbook.'

    $a1NavigationCount = ([regex]::Matches($b, 'File\.Contents\([^\r\n]*CapacityDistrib\(A\.1\)-shifts\.xlsx')).Count
    Add-ContractResult -Name 'merged-b-single-a1-navigation' -File 'B' -Passed ($a1NavigationCount -eq 1) -Detail "Expected one A1 File.Contents navigation; found $a1NavigationCount."
    $bA1Source = Add-DefinitionPresence -Stage 'B' -Text $b -Name 'IMPORTSource A1'
    if ($null -ne $bA1Source) {
        $compact = Get-CompactM $bA1Source
        Add-ContractResult -Name 'merged-b-buffered-a1-source' -File 'B' -Passed (
            $compact.Contains('Binary.Buffer(File.Contents(') -and $compact.Contains('Excel.Workbook(')
        ) -Detail 'Merged B must create one buffered A1 binary/navigation source for all internal imports.'
    }

    $bHelperNames = @(Get-MDeclarationNames -Text $b | Where-Object { $_ -eq 'fnAddIndexedRunningTotal' })
    Add-ContractResult -Name 'merged-b-single-running-helper' -File 'B' -Passed ($bHelperNames.Count -eq 1) -Detail "Expected one shared running-total helper; found $($bHelperNames.Count)."

    $bContract = Add-DefinitionPresence -Stage 'B' -Text $b -Name 'IMPORT ResourceContract'
    if ($null -ne $bContract) {
        Add-ContractResult -Name 'b-contract-source-role' -File 'B' -Passed ((Get-CompactM $bContract).Contains('SourceRole')) -Detail 'External Resource contracts must filter the upstream AIN role with SourceRole.'
    }
    $bCapped = Get-MSharedDefinition -Text $b -Name 'ResPeriodAvailabilityCapped(C#)TABLE'
    if ($null -ne $bCapped) {
        $compact = Get-CompactM $bCapped
        Add-ContractResult -Name 'b-capped-output-role' -File 'B' -Passed ($compact.Contains('OutputRole')) -Detail 'Internal/published C# rows must filter or label the experimental OutputRole.'
        Add-ContractResult -Name 'merged-b-capped-validation-gates' -File 'B' -Passed (
            $compact.Contains('RoleContext_CHECK') -and $compact.Contains('A2_COMPLETE_GRID_CHECK') -and $compact.Contains('ResPeriodCappedAvailability(C#)TABLE')
        ) -Detail 'The published compatibility C# interface must force role and complete-grid validation before consuming the internal stage.'
    }

    foreach ($spec in @(
        @{ Name = 'ResCapPrioritisedTABLE'; Key = 'Period' },
        @{ Name = 'PeriodCapPrioritisedTABLE'; Key = 'Resource' },
        @{ Name = 'PrioritiseRedistribAvail-Distrib'; Key = 'Period' },
        @{ Name = 'PrioritiseReductionAvail-Setup'; Key = 'Resource' },
        @{ Name = 'ResIndex-Steup'; Key = 'Period' }
    )) {
        $definition = Add-DefinitionPresence -Stage 'B' -Text $b -Name $spec.Name
        if ($null -ne $definition) {
            Add-ContractResult -Name "stable-sort:$($spec.Name)" -File 'B' -Passed (Test-MSortContainsKey -Definition $definition -ExpectedKey $spec.Key) -Detail "The merged priority sort must include $($spec.Key) as a deterministic key."
        }
    }

    $a2A1Source = Add-DefinitionPresence -Stage 'A2' -Text $a2 -Name 'IMPORTSource A1'
    if ($null -ne $a2A1Source) {
        $compact = Get-CompactM $a2A1Source
        Add-ContractResult -Name 'a2-single-buffered-a1-source' -File 'A2' -Passed (
            $compact.Contains('Binary.Buffer(File.Contents(') -and $compact.Contains('Excel.Workbook(') -and
            ([regex]::Matches($a2, [regex]::Escape('CapacityDistrib(A.1)-shifts.xlsx'))).Count -eq 1
        ) -Detail 'The improved three-file A2 reference must use one buffered A1 source.'
    }
}

$nativeResult = $null
if ($RunNative) {
    $nativeScript = Join-Path $PSScriptRoot 'test-capacity-distribution-ain2-native.ps1'
    try {
        $nativeOutput = & pwsh -NoProfile -File $nativeScript 2>&1
        $nativeExitCode = $LASTEXITCODE
        $nativeResult = ($nativeOutput -join "`n")
        Add-ContractResult -Name 'native-synthetic-fixtures' -File 'Diagnostics' -Passed ($nativeExitCode -eq 0) -Detail $nativeResult
    }
    catch {
        Add-ContractResult -Name 'native-synthetic-fixtures' -File 'Diagnostics' -Passed $false -Detail $_.Exception.Message
    }
}

$failures = @($results | Where-Object { -not $_.Passed })
$summary = [ordered]@{
    Status = if ($failures.Count -eq 0) { 'Passed' } else { 'Failed' }
    Checks = $results.Count
    Failed = $failures.Count
    SourceRoot = $SourceRoot
    NativeSkipped = -not [bool]$RunNative
    Failures = $failures
}
$summary | ConvertTo-Json -Depth 8
if ($failures.Count -gt 0) { exit 1 }
