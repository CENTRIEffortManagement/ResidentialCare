[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
Set-Location -LiteralPath $repoRoot
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $repoRoot 'outputs/te-availability-role-fix/sync-powerquery.ps1'), [ref] $tokens, [ref] $errors)
if ($errors.Count) { throw 'Approved M definition reader did not parse.' }
foreach ($name in @('Read-MDefinitions', 'Normalize-Formula')) {
    $definition = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false) |
        Where-Object Name -eq $name)
    if ($definition.Count -ne 1) { throw "Missing definition reader: $name" }
    . ([scriptblock]::Create($definition[0].Extent.Text))
}
$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'archive-manifest.json') -Raw | ConvertFrom-Json
$comparisons = foreach ($item in $manifest.Archives) {
    if ((Get-FileHash -LiteralPath $item.Archive).Hash -cne $item.ArchiveSHA256) { throw 'Archive changed.' }
    $workbookName = [IO.Path]::GetFileNameWithoutExtension($item.Workbook)
    $evidence = Get-Content -LiteralPath (Join-Path $PSScriptRoot "verification/$workbookName.extraction.json") -Raw | ConvertFrom-Json
    if ((Get-FileHash -LiteralPath $item.Workbook).Hash -cne $evidence.workbookSHA256) { throw 'Workbook changed during source-only work.' }
    $before = Read-MDefinitions $item.Archive
    $after = Read-MDefinitions $item.Source
    $added = @($after.Keys | Where-Object { -not $before.Contains($_) })
    $removed = @($before.Keys | Where-Object { -not $after.Contains($_) })
    $changed = @($after.Keys | Where-Object {
        $before.Contains($_) -and (Normalize-Formula $before[$_]) -cne (Normalize-Formula $after[$_])
    })
    $isCapacity = $workbookName -eq 'Capacity-ShiftAvailability'
    if ($isCapacity) {
        $expectedAdded = @('IMPORT Published Roster', 'PublishedRoster_Prepare', 'RosterWorkers_Prepare',
            'WorkerEmployment_Prepare', 'WorkerEmployment_Selected', 'AvailabilityMembership_Prepare')
        $expectedChanged = @("IMPORT Worker's Report", 'ReconciledWorkerMembership_Prepare', 'WorkerContracts_Prepare',
            'ReconciledWorkers_Prepare', 'ReconciledWorkers_Eligible', 'ReconciledWorkers_CHECK')
        if (@(Compare-Object $expectedAdded $added -CaseSensitive).Count -or
            $removed.Count -ne 1 -or $removed[0] -cne 'IMPORT Reconciled Workers' -or
            @(Compare-Object $expectedChanged $changed -CaseSensitive).Count) { throw 'Unexpected query differences.' }
        $text = [IO.File]::ReadAllText($item.Source)
        if ($text -match 'Worker Reconciliation\.xlsx|Leave balances only|#"IMPORT Reconciled Workers"') { throw 'Obsolete dependency or leave-only branch remains.' }
        foreach ($name in @('Availability-StaffList', 'ResDayShift', 'RoleAvailabilityCapped',
            'RoleResDayAvailabilityCapped', 'ResDayShift_CHECK', 'AllocationRoleShares_CHECK', 'AvailabilityRules_TEST',
            'ReconciledWorkers_Facility_Prepare', 'FacilityMappings_Prepare')) {
            if ((Normalize-Formula $before[$name]) -cne (Normalize-Formula $after[$name])) { throw "Required unchanged interface or calculation changed: $name" }
        }
        if ($after['ReconciledWorkerMembership_Prepare'] -notmatch '"In Leave Balances", each false, type logical') {
            throw 'Compatibility Balance flag is not explicitly false.'
        }
        $newImport = $after['IMPORT Published Roster']
        $workerImport = $after["IMPORT Worker's Report"]
        if ($newImport -notmatch 'Allocation\\Published Roster\.xlsx' -or $newImport -notmatch 'Item = "Roster Data"' -or
            ($newImport + $workerImport) -match 'Table\.(Select|Transform|Promote|Rename|Add|Remove|Group|Join|NestedJoin|Replace|Sort|Distinct)') {
            throw 'New or materially edited imports must land approved source objects unchanged.'
        }
    } elseif ($added.Count -or $removed.Count -or $changed.Count -or (Get-FileHash -LiteralPath $item.Source).Hash -cne $item.SHA256) {
        throw 'StaffListMaster must remain byte-identical to its approved baseline.'
    }
    [ordered]@{Source=$item.Source; Workbook=$item.Workbook; Archive=$item.Archive;
        BaselineQueryCount=$before.Count; RevisedQueryCount=$after.Count; Added=$added; Removed=$removed; Changed=$changed;
        SourceSHA256=(Get-FileHash -LiteralPath $item.Source).Hash; ArchiveSHA256=$item.ArchiveSHA256;
        WorkbookSHA256=$evidence.workbookSHA256; WorkbookUnchanged=$true}
}
$result = [ordered]@{Status='Passed'; Scope='BD originals only; approved workbook-derived baseline';
    Comparisons=@($comparisons); PublishedInterfacesUnchanged=$true; ExistingFacilityMappingPreserved=$true;
    BalanceWorkbookDependencyRemoved=$true; ArchivesVerified=$true; Synced=$false; Refreshed=$false;
    RuntimeAcceptance='Pending manual workbook sync and evaluated checks'}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'source-validation.json') -Encoding utf8
$result | ConvertTo-Json -Depth 8
