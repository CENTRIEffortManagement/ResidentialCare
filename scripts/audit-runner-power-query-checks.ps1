param(
    [string] $OutputDirectory = 'outputs/runner-check-audit'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$planningScript = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/runner/src/BatchPlanning.ps1'
. $planningScript

function ConvertFrom-MIdentifier {
    param([Parameter(Mandatory)] [string] $Identifier)

    $trimmed = $Identifier.Trim()
    if ($trimmed.StartsWith('#"') -and $trimmed.EndsWith('"')) {
        return $trimmed.Substring(2, $trimmed.Length - 3).Replace('""', '"')
    }
    return $trimmed
}

function Get-MQueryBlocks {
    param([Parameter(Mandatory)] [string] $Path)

    $text = [IO.File]::ReadAllText($Path)
    $queryMatches = [regex]::Matches($text, '(?m)^\s*shared\s+(?<name>#"(?:[^"]|"")*"|[^=\r\n]+?)\s*=')
    $queries = [Collections.Generic.List[object]]::new()

    for ($index = 0; $index -lt $queryMatches.Count; $index++) {
        $start = $queryMatches[$index].Index
        $end = if (($index + 1) -lt $queryMatches.Count) { $queryMatches[$index + 1].Index } else { $text.Length }
        $rawName = $queryMatches[$index].Groups['name'].Value.Trim()
        $block = $text.Substring($start, $end - $start)
        $code = (($block -split "`r?`n") | Where-Object { $_ -notmatch '^\s*//' }) -join "`n"
        $name = ConvertFrom-MIdentifier $rawName

        $queries.Add([pscustomobject]@{
            Name = $name
            RawName = $rawName
            Block = $block
            IsCheckLike = [bool] ($name -match '(?i)(check|validat|reconcil|diagnos|issue|audit|test|error|exception|mismatch|missing|duplicate|integrity|quality)')
            HasHardError = [bool] ($code -match '(?i)(^|[^A-Za-z])error\s+(Error\.Record|"|\()')
            CapturesError = [bool] ($code -match '(?is)\btry\b.+\botherwise\b|\[HasError\]')
            HasEmptyShortCircuit = [bool] ($code -match '(?is)\bif\b.{0,240}\b(Table|List)\.IsEmpty\b|\bif\b.{0,240}\bTable\.RowCount\s*\([^\)]*\)\s*(=|<>|<|>)')
            Consumers = @()
        })
    }

    foreach ($query in $queries) {
        $token = if ($query.RawName.StartsWith('#"')) {
            [regex]::Escape($query.RawName)
        }
        else {
            '(?<![A-Za-z0-9_])' + [regex]::Escape($query.RawName) + '(?![A-Za-z0-9_])'
        }
        $query.Consumers = @($queries | Where-Object {
                $_.Name -ne $query.Name -and $_.Block -match $token
            } | ForEach-Object Name)
    }

    return @($queries)
}

function ConvertTo-MarkdownCell {
    param([AllowEmptyString()] [string] $Value)
    return (($Value -replace '\|', '\|') -replace "`r?`n", '<br>')
}

$catalogue = Get-BatchCatalogue -RepoRoot $repoRoot
$defaultJobIds = @{}
foreach ($job in @((Select-BatchPlan $catalogue -RunAll).Jobs)) { $defaultJobIds[$job.Id] = $true }
$roleGroups = @($catalogue.Jobs | Where-Object { $_.Role } | Group-Object Unit, Role)
$representativeRoleGroup = @($roleGroups | Where-Object {
        $_.Count -eq 3 -and
        $defaultJobIds.ContainsKey($_.Group[0].Id) -and
        @($_.Group | Where-Object { Test-Path -LiteralPath ($_.Path + '_PowerQuery.m') -PathType Leaf }).Count -eq 3
    } | Select-Object -First 1)
if ($representativeRoleGroup.Count -ne 1) {
    throw 'No complete source-backed three-workbook role folder exists in the default run for representative audit.'
}
$representativeRoleJobIds = @{}
foreach ($job in $representativeRoleGroup[0].Group) { $representativeRoleJobIds[$job.Id] = $true }
$representativeRole = "$($representativeRoleGroup[0].Group[0].Unit)/$($representativeRoleGroup[0].Group[0].Role)"
$auditJobs = @($catalogue.Jobs | Where-Object {
        -not $_.Role -or $representativeRoleJobIds.ContainsKey($_.Id)
    })
$omittedEquivalentRoleJobs = @($catalogue.Jobs | Where-Object {
        $_.Role -and -not $representativeRoleJobIds.ContainsKey($_.Id)
    })

$fileRows = [Collections.Generic.List[object]]::new()
$queryRows = [Collections.Generic.List[object]]::new()

foreach ($job in $auditJobs) {
    $sourcePath = $job.Path + '_PowerQuery.m'
    $sourceRelativePath = $sourcePath.Substring($repoRoot.Length + 1).Replace('\', '/')
    $inDefaultRun = $defaultJobIds.ContainsKey($job.Id)

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        $fileRows.Add([pscustomobject]@{
            Job = $job.Id
            Workbook = $job.RelativePath
            InDefaultRun = $inDefaultRun
            SourceStatus = 'Missing exact extracted source'
            Source = $sourceRelativePath
            QueryCount = 0
            NamedCheckCount = 0
            HardGuardCount = 0
            ErrorCaptureCount = 0
            EmptyShortCircuitCount = 0
            UnreferencedCheckCount = 0
            NamedChecks = ''
        })
        continue
    }

    $queries = @(Get-MQueryBlocks $sourcePath)
    $namedChecks = @($queries | Where-Object { $_.IsCheckLike })
    $hardGuards = @($queries | Where-Object { $_.HasHardError })
    $errorCaptures = @($queries | Where-Object { $_.CapturesError })
    $shortCircuits = @($queries | Where-Object { $_.HasEmptyShortCircuit })
    $unreferencedChecks = @($namedChecks | Where-Object { $_.Consumers.Count -eq 0 })

    $fileRows.Add([pscustomobject]@{
        Job = $job.Id
        Workbook = $job.RelativePath
        InDefaultRun = $inDefaultRun
        SourceStatus = 'Audited'
        Source = $sourceRelativePath
        QueryCount = $queries.Count
        NamedCheckCount = $namedChecks.Count
        HardGuardCount = $hardGuards.Count
        ErrorCaptureCount = $errorCaptures.Count
        EmptyShortCircuitCount = $shortCircuits.Count
        UnreferencedCheckCount = $unreferencedChecks.Count
        NamedChecks = (@($namedChecks | ForEach-Object Name) -join ' | ')
    })

    foreach ($query in @($queries | Where-Object { $_.IsCheckLike -or $_.HasHardError -or $_.CapturesError -or $_.HasEmptyShortCircuit })) {
        $classification = [Collections.Generic.List[string]]::new()
        if ($query.IsCheckLike) { $classification.Add('Named check or diagnostic') }
        if ($query.HasHardError) { $classification.Add('Raises M error') }
        if ($query.CapturesError) { $classification.Add('Captures or replaces error') }
        if ($query.HasEmptyShortCircuit) { $classification.Add('Empty/count short-circuit') }

        $runnerRisk = if ($query.IsCheckLike -and $query.Consumers.Count -eq 0) {
            'Not demanded by another source query; runner cannot assume it executed or inspect its result'
        }
        elseif ($query.HasHardError) {
            'Can block a calculation, but current runner only waits for idle/calculation state'
        }
        elseif ($query.CapturesError) {
            'Can convert a calculation error into data, null, fallback, warning, or empty output'
        }
        elseif ($query.HasEmptyShortCircuit) {
            'Can publish an empty or alternate result without a runner-visible failure'
        }
        else {
            'Diagnostic result is not interpreted by the runner'
        }

        $queryRows.Add([pscustomobject]@{
            Job = $job.Id
            Workbook = $job.RelativePath
            Source = $sourceRelativePath
            Query = $query.Name
            InDefaultRun = $inDefaultRun
            Classification = ($classification -join '; ')
            ConsumerCount = $query.Consumers.Count
            Consumers = ($query.Consumers -join ' | ')
            RunnerRisk = $runnerRisk
        })
    }
}

$outputPath = if ([IO.Path]::IsPathRooted($OutputDirectory)) {
    [IO.Path]::GetFullPath($OutputDirectory)
}
else {
    [IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDirectory))
}
if (-not $outputPath.StartsWith($repoRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputDirectory must resolve inside the repository.'
}
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

$filesCsv = Join-Path $outputPath 'workbook-check-inventory.csv'
$queriesCsv = Join-Path $outputPath 'query-check-inventory.csv'
$summaryPath = Join-Path $outputPath 'README.md'
$fileRows | Export-Csv -LiteralPath $filesCsv -NoTypeInformation -Encoding utf8
$queryRows | Export-Csv -LiteralPath $queriesCsv -NoTypeInformation -Encoding utf8

$auditedFiles = @($fileRows | Where-Object { $_.SourceStatus -eq 'Audited' })
$missingFiles = @($fileRows | Where-Object { $_.SourceStatus -ne 'Audited' })
$filesWithNamedChecks = @($auditedFiles | Where-Object { $_.NamedCheckCount -gt 0 })
$filesWithHardGuards = @($auditedFiles | Where-Object { $_.HardGuardCount -gt 0 })
$unreferencedCheckQueries = @($queryRows | Where-Object {
        $_.Classification -like '*Named check or diagnostic*' -and $_.ConsumerCount -eq 0
    })
$defaultFiles = @($fileRows | Where-Object { $_.InDefaultRun })
$defaultMissing = @($defaultFiles | Where-Object { $_.SourceStatus -ne 'Audited' })
$batchLogRoot = Join-Path $repoRoot 'CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh'
$latestRun = if (Test-Path -LiteralPath $batchLogRoot -PathType Container) {
    Get-ChildItem -LiteralPath $batchLogRoot -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}
else { $null }
$latestRefreshLogs = if ($null -ne $latestRun) { @(Get-ChildItem -LiteralPath $latestRun.FullName -Recurse -Filter 'refresh.log' -File) } else { @() }
$checkFlagPatterns = '(?i)(query check|check result|check failed|validation result|power query error|cell error|formula error)'
$latestCheckFlags = @($latestRefreshLogs | Select-String -Pattern $checkFlagPatterns)
$zeroQueryTableLogs = @($latestRefreshLogs | Where-Object {
        (Get-Content -LiteralPath $_.FullName -Raw) -match 'supported refresh object\(s\): \d+ connection\(s\), 0 worksheet query table\(s\)'
    })

$lines = [Collections.Generic.List[string]]::new()
$lines.Add('# ResidentialCare runner check audit')
$lines.Add('')
$lines.Add('This is a source-only audit of the non-role workbooks in the current batch catalogue plus one representative three-workbook role folder. It does not inspect or alter an Excel workbook.')
$lines.Add('')
$lines.Add('## Result')
$lines.Add('')
$lines.Add("- Catalogue workbooks: $($catalogue.Jobs.Count)")
$lines.Add("- Workbooks in this audit: $($fileRows.Count)")
$lines.Add("- Representative role folder: $representativeRole (three files)")
$lines.Add("- Equivalent role-folder workbook instances omitted: $($omittedEquivalentRoleJobs.Count)")
$lines.Add("- Exact extracted M sources audited: $($auditedFiles.Count)")
$lines.Add("- Workbooks blocked from source audit by missing exact extracted M: $($missingFiles.Count)")
$lines.Add("- Default-run workbooks in this audit: $($defaultFiles.Count); source-backed: $($defaultFiles.Count - $defaultMissing.Count); missing source: $($defaultMissing.Count)")
$lines.Add("- Source-backed workbooks with named checks/validation/diagnostics: $($filesWithNamedChecks.Count)")
$lines.Add("- Named check-like queries: $(($auditedFiles | Measure-Object NamedCheckCount -Sum).Sum)")
$lines.Add("- Source-backed workbooks with explicit M error gates: $($filesWithHardGuards.Count)")
$lines.Add("- Queries containing explicit M error gates: $(($auditedFiles | Measure-Object HardGuardCount -Sum).Sum)")
$lines.Add("- Named check-like queries with no source-query consumer: $($unreferencedCheckQueries.Count)")
$lines.Add('')
$lines.Add('## Runner finding')
$lines.Add('')
$lines.Add('The expanded runner currently treats refresh-idle plus Excel calculation-done as success. It does not enumerate Power Query checks, interpret returned check tables, scan loaded outputs for Excel errors, or prove that connection-only check queries were demanded. It can therefore save and report success without naming a failed, skipped, or unobserved workbook check.')
$lines.Add('')
$lines.Add('A safe fail-closed contract needs every required check to have a machine-readable result and an inspectable load surface. A query name alone is not enough: many check queries return tables, some raise M errors, some capture upstream errors, and some are not referenced by another source query.')
$lines.Add('')
$lines.Add('## Observed runner evidence')
$lines.Add('')
if ($null -eq $latestRun) {
    $lines.Add('- No batch-run log directory was available.')
}
else {
    $latestRunRelative = $latestRun.FullName.Substring($repoRoot.Length + 1).Replace('\', '/')
    $lines.Add(('- Latest run inspected: `{0}`' -f $latestRunRelative))
    $lines.Add("- Workbook refresh logs: $($latestRefreshLogs.Count)")
    $lines.Add("- Explicit workbook-query/check-result flag lines: $($latestCheckFlags.Count)")
    $lines.Add("- Logs reporting zero worksheet query tables: $($zeroQueryTableLogs.Count)")
}
$lines.Add('')
$lines.Add('## Required fail-closed runner contract')
$lines.Add('')
$lines.Add('1. After refresh and calculation settle, but before Save, evaluate every required workbook check through a declared, machine-readable result surface.')
$lines.Add('2. Treat a failed check, an Excel/PQ error, a required check that did not run, or a required check that cannot be inspected as a workbook failure. Log the workbook, query/check name, status, count/value, and message.')
$lines.Add('3. Close the failed workbook without saving, stop new batch dispatch, and mark every not-started job blocked by the check failure. Do not continue unrelated branches after a calculation-integrity failure.')
$lines.Add('4. Keep advisory diagnostics separate from required gates. Names such as `DIAGNOSTICS` and `TEST` cannot automatically be treated as failures without a declared policy.')
$lines.Add('5. Extract the missing exact M sources before changing checks in those workbooks. Do not infer their contract from a copy, backup, donor, or same-stem file.')
$lines.Add('')
$lines.Add('## Missing exact sources')
$lines.Add('')
$lines.Add('| Default | Job | Workbook | Expected source |')
$lines.Add('|---|---|---|---|')
foreach ($row in $missingFiles) {
    $lines.Add("| $(if ($row.InDefaultRun) { 'Yes' } else { 'No' }) | $(ConvertTo-MarkdownCell $row.Job) | $(ConvertTo-MarkdownCell $row.Workbook) | $(ConvertTo-MarkdownCell $row.Source) |")
}
$lines.Add('')
$lines.Add('## Named checks with no source-query consumer')
$lines.Add('')
$lines.Add('These checks are not demanded by another query in the extracted source. They may be loaded independently, but the current runner neither proves that nor interprets their result.')
$lines.Add('')
$lines.Add('| Default | Workbook | Query | Raises M error |')
$lines.Add('|---|---|---|---|')
foreach ($row in @($unreferencedCheckQueries | Sort-Object Workbook, Query)) {
    $lines.Add("| $(if ($row.InDefaultRun) { 'Yes' } else { 'No' }) | $(ConvertTo-MarkdownCell $row.Workbook) | $(ConvertTo-MarkdownCell $row.Query) | $(if ($row.Classification -like '*Raises M error*') { 'Yes' } else { 'No' }) |")
}
$lines.Add('')
$lines.Add('## Source-backed files with checks or gates')
$lines.Add('')
$lines.Add('| Default | Workbook | Named checks | Hard gates | Captures | Empty/count branches | Unreferenced checks |')
$lines.Add('|---|---|---:|---:|---:|---:|---:|')
foreach ($row in @($auditedFiles | Where-Object {
            $_.NamedCheckCount -gt 0 -or $_.HardGuardCount -gt 0 -or $_.ErrorCaptureCount -gt 0 -or $_.EmptyShortCircuitCount -gt 0
        } | Sort-Object Workbook)) {
    $lines.Add("| $(if ($row.InDefaultRun) { 'Yes' } else { 'No' }) | $(ConvertTo-MarkdownCell $row.Workbook) | $($row.NamedCheckCount) | $($row.HardGuardCount) | $($row.ErrorCaptureCount) | $($row.EmptyShortCircuitCount) | $($row.UnreferencedCheckCount) |")
}
$lines.Add('')
$lines.Add('## Detailed inventories')
$lines.Add('')
$lines.Add('- `workbook-check-inventory.csv` contains every catalogue workbook, including missing-source blockers.')
$lines.Add('- `query-check-inventory.csv` contains every query with a check-like name, explicit M error, error capture/fallback, or empty/count branch.')
$lines.Add('')
$lines.Add('Counts are static-source indicators, not proof that a condition fails at runtime. Workbook load settings and current query results remain unverified unless they are exposed through an approved source and a runner-readable contract.')

[IO.File]::WriteAllLines($summaryPath, $lines, [Text.UTF8Encoding]::new($false))

Write-Host "Audit complete: $($fileRows.Count) selected workbook(s), $($auditedFiles.Count) source-backed, $($missingFiles.Count) missing exact source."
Write-Host "Named check-like queries: $(($auditedFiles | Measure-Object NamedCheckCount -Sum).Sum); explicit M-error queries: $(($auditedFiles | Measure-Object HardGuardCount -Sum).Sum); unreferenced named checks: $($unreferencedCheckQueries.Count)."
Write-Host "Report: $($summaryPath.Substring($repoRoot.Length + 1))"
