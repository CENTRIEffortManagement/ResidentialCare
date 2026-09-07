# MinuteWorker FTE Analysis — PC Handoff

**Snapshot date:** 30 August 2026, Australia/Sydney.

**Revision:** 7 September 2026, whole-period distribution revision. `TargetMinutes` contains productive-care **hours per fortnight**. The final result is required roster FTE; care minutes are an intermediate calculation. The allocator now preserves differences between weekdays using whole-period historical shares, then displays the average of corresponding weekdays. Earlier daily-minute input assumptions and fixed-category-target-per-weekday rules are superseded.

> **Staleness warning:** This document describes the work as it stood on the date above. If Codex or another developer reads it much later, it may be out of date. Before making changes, re-read the repository `AGENTS.md`, check the current Git/worktree state, inspect the current `.m` source, and verify the workbook table structures and input values.

## Goal

Produce a MinuteWorker-only analysis of required roster FTE by facility, role, weekday and shift, using funded productive-care hours per fortnight and whole-period historical roster patterns. Retain unaveraged historical FTE by original roster week to compare the two Mondays, two Tuesdays, and so on before accepting a representative-week pattern.

The analysis is a separate branch. Existing roster-analysis outputs should remain unchanged.

## Current working files

- Power Query source: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m`
- Target workbook: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx`
- This handoff: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/MinuteWorker-FTE-Handoff-2026-08-30.md`

During this work, the user explicitly instructed Codex to use the current `.m` file and the approved **IMPORT Extract** method. Do not overwrite the current `.m` by extracting older code from the workbook. Confirm the source-of-truth state before any future extraction or synchronization. Repository policy normally places canonical product-specific Power Query source under `Workflows/`, so reconcile that policy with the explicitly selected sidecar source before moving or syncing anything.

The whole-period distribution revision and new historical table were synchronized through the approved Excel Power Query Editor extension on 7 September at approximately 20:19 Australia/Sydney. The extension reported success and created `Demand-MasterRoster Manual Read.xlsx.backup.2026-09-07T10-19-34-585Z` beside the workbook. Automatic backup cleanup was temporarily disabled for this sync; the temporary settings overrides were then removed without changing the existing settings. Excel runtime refresh/validation and loading the new historical output to a worksheet are still pending. Earlier refreshes and observed PASS values do not validate this revision. An earlier accidental `MW Distribution Check` header edit was restored. Preserve the workbook's existing worktree changes; do not extract older workbook code over this source.

## Confirmed business rules

1. `INPUT MinuteWorkers` is the role list and contains each role's `Direct Care %`.
2. Despite its name, `INPUT TargetMinutes` contains productive-care **hours per fortnight (14 days)**, confirmed by the user on 7 September 2026.
3. Targets are supplied for each facility for two types:
   - `RN`
   - `ALL`, which includes RN hours
4. The calculated MinuteCategories are:
   - `RN`: the canonical role `RN` only
   - `OTHERS`: every other MinuteWorker role, including RN-qualified management roles
5. Average daily category targets (reference values, not a fixed target for every weekday) are:
   - `RNDailyTargetMinutes = RNFortnightTargetHours * 60 / 14`
   - `OTHERSDailyTargetMinutes = (ALLFortnightTargetHours - RNFortnightTargetHours) * 60 / 14`
6. Fortnight category minute targets equal the input category hours multiplied by 60. Original `RNFortnightTargetHours` and `ALLFortnightTargetHours` remain visible in target preparation.
7. Historical hours are filtered to MinuteWorker roles and adjusted for direct care:
   - `HistoricalProductiveHours = HistoricalRosterHours * Direct Care %`
8. Retain each original `Week No` at facility/role/weekday/shift grain. Average corresponding cells across all observed facility weeks, including zero cells in complete weeks. With two weeks, this is `(Week 1 Monday + Week 2 Monday) / 2`, not the average of only populated cells. Incomplete source weeks fail validation.
9. The conditional role share within a weekday remains available as a diagnostic:
   - `RoleDayHistoryDistribution% = RoleDayHistoricalProductiveHours / CategoryDayHistoricalProductiveHours`
10. The conditional shift share within a role/weekday also remains available:
   - `RoleDayShiftDistribution% = HistoricalRosterHours / RoleDayHistoricalRosterHours`
11. Allocate the whole-period category budget across roles, weekdays and shifts:
    - `CategoryWeeklyHistoricalProductiveHours = sum(HistoricalProductiveHours)` across all representative-week cells in the facility/category.
    - `CategoryWeekRoleDayShiftDistribution% = HistoricalProductiveHours / CategoryWeeklyHistoricalProductiveHours`.
    - `CategoryWeekdayDistribution% = CategoryDayHistoricalProductiveHours / CategoryWeeklyHistoricalProductiveHours`.
    - `CategoryWeekdayTargetMinutes = CategoryTargetMinutes / 2 * CategoryWeekdayDistribution%`.
    - `RoleDailyTargetMinutes = CategoryWeekdayTargetMinutes * RoleDayHistoryDistribution%`.
    - `WeekdayShiftTargetMinutes = CategoryTargetMinutes / 2 * CategoryWeekRoleDayShiftDistribution%`.
    - equivalently, `WeekdayShiftTargetMinutes = RoleDailyTargetMinutes * RoleDayShiftDistribution%`
12. Productive-care target minutes are converted back to roster minutes:
    - `WeekdayShiftRosterMinutes = WeekdayShiftTargetMinutes / Direct Care %`
13. One FTE shift is 7.6 hours or 456 minutes:
    - `FTE = WeekdayShiftRosterMinutes / 456`
    - `FTE` means 7.6-hour shift equivalents, not weekly employee FTE. Sum the three shifts for a daily total. `RoleAverageDailyTargetMinutes` is a repeated informational field and must not be summed across shift rows.
14. Calculations retain full precision. Rounding is for display only.
15. A positive facility/category target with no eligible productive history across the entire period fails validation. A zero-weight category/weekday is allowed if source coverage is complete; no allocation is invented for it.
16. Weekdays compete for the same finite category budget. Changing Tuesday history can therefore change Monday's absolute target FTE, although Monday's conditional role/shift mix is unchanged. Busier Tuesdays are preserved rather than flattened to the daily average.
17. Normalizing seven-day average history gives the same relative weights as pooling corresponding weekday cells across the source weeks. Allocate half the fortnight budget to that representative week; double its totals only for fortnight reconciliation. Do not divide the resulting FTE by two again.
18. `MW Fortnight Hours Check` independently compares `sum(FTE * 7.6 * Direct Care % * 2)` over the representative week with the original RN and ALL input hours. Failures block TABLE and MATRIX through `MW Publication Check`. This comparison reads raw hours directly rather than reusing converted minute targets.
19. Historical FTE is `HistoricalRosterHours / 7.6`, before target scaling and without dividing by Direct Care %. Retain original week identifiers for chart series. If more than two weeks are supplied, preserve all of them and expose `HistoricalWeeksInAverage`; do not silently select two.

## Role normalization

`MW Abbreviate Role` is applied both to the raw `IMPORT Master` roles and to `INPUT MinuteWorkers`. `Master Prepare` filters dynamically using the canonical MinuteWorker role list instead of an incomplete hard-coded list.

| Source role variants | Canonical role |
| --- | --- |
| Registered Nurse; REGN variants | RN |
| Assistant/Asst in Nursing; AINC4/Med Comp variants | AIN |
| Enrolled Nurse | EN |
| Clinical Care Coordinator; Clinical Care Coordinator & Manager | CCCM |
| Residential Services Manager | RSM |
| Care Service Manager; Care Services Manager | CSM |
| Wellbeing & Care Support Officer; Wellbeing & Lifestyle Officer; WLO | WCSO |
| Care & Assessment Manager | CAM |
| Regional General Manager | RGM |
| Therapy Assistant | TA |

The workbook's MinuteWorker list was last observed as RN, AIN, EN, CCCM, RSM, CSM, WCSO, CAM, and RGM. CAM and RGM did not have matching historical roster hours in the data inspected on the snapshot date. They should remain visible as warnings/unmatched roles; no historical allocation is to be invented for them.

## Query dependency and presentation order

The `.m` file presents substantial steps as separate named queries with `// Query:` and `// Purpose:` headers and business-rule comments.

### Inputs and preparation

- `IMPORT Master`
- `INPUT MinuteWorkers`
- `MW Abbreviate Role`
- `MW MinuteWorkers Prepare`
- `Master Prepare`
- `LocRoleWeekDaysHours` (legacy aggregation and source-key/coverage checks)
- existing legacy historical queries ending in `LocRoleDayShift%` remain for non-MinuteWorker outputs only
- `INPUT TargetMinutes`
- `MW TargetMinutes Prepare`

### Historical distribution

- `MW Historical WeekDayShift` (unaveraged complete-cell staging, directly from `Master Prepare`)
- `MW Historical DayShift`
- `MW Role Distribution`

### Allocation

- `MW Category Weekday Targets`
- `MW Role Targets`
- `MW DayShift Allocation`

### Validation

- `MW Input Check`
- `MW Distribution Check`
- `MW Daily Allocation Check`
- `MW PreAllocation Check`
- `MW Fortnight Hours Check`
- `MW Publication Check`

### Published outputs

- `MinuteWorkerRoleAssignments_TABLE`
- `MinuteWorkersFTE_TABLE`
- `MinuteWorkersFTE_MATRIX`
- `MinuteWorkersFTE_CHECK`
- `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK`
- `MinuteWorkersFTE_CATEGORY_DAILY_CHECK`
- `MinuteWorkersFTE_HISTORICAL_DAY_CHECK`
- `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE`

Staging queries should be connection-only in Excel when supported. Load only the outputs that users need to worksheets or the data model.

## Output grains and purpose

### `MinuteWorkerRoleAssignments_TABLE`

One row per exact original Master Roster role and assigned configured MinuteWorker role. It exposes `MinuteCategory`, `QFR Category`, `Direct Care %`, source-row count, facility count, and facility list. Unmatched Master Roster roles are deliberately excluded.

### `MinuteWorkersFTE_TABLE`

Detailed allocation at approximately:

`Facility × MinuteCategory × Role × Weekday × Shift`

It contains the historical measures and distributions, direct-care percentage, daily/fortnight/role targets, allocated productive minutes, roster minutes, and unrounded FTE.

### `MinuteWorkersFTE_MATRIX`

One row per facility/category/role, with weekday-shift combinations as FTE columns. This is the user-facing staffing pattern view.

### `MinuteWorkersFTE_CHECK`

The principal validation output. It surfaces input, matching, distribution, allocation, and reconciliation failures, including independent checks against original fortnight hours. For those checks, Actual and Expected are in hours; existing minute-reconciliation checks remain in minutes. Required publication checks must pass before using TABLE or MATRIX.

### `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK`

The role proof table. It has seven rows for every facility/category/role, including zero-allocation days, and proves that each role/day target, the representative week, and the reconstructed fortnight reconcile.

Important columns include:

- `RoleDayHistoryDistribution%`
- `DailyShiftDistributionTotal%`
- `RoleDailyTargetMinutes`
- `DayAllocationVarianceMinutes`
- `DayVsRoleDayTarget%`
- `MaximumAbsoluteTwoStageVarianceMinutes`
- `AllocatedDayProductiveMinutes`
- `AllocatedDayRosterMinutes`
- `DayFTEShiftTotal`
- `ExpectedWeeklyProductiveMinutes`
- `AllocatedWeeklyProductiveMinutes`
- `WeeklyVarianceMinutes`
- `RoleHistoryDistributionAcrossWeek%` (informational sum of seven day-specific shares; it is not expected to equal 100%)
- `DaysWithAllocatedMinutes`
- `AllocatedWeeklyRosterMinutes`
- `WeeklyFTEShiftTotal`
- `ReconstructedFortnightProductiveMinutes`
- `FortnightVarianceMinutes`
- `DayVsAverageDailyTarget%`
- `CumulativeWeekProductiveMinutes`
- `CumulativeWeeklyTarget%`
- `Status`
- `CheckMessage`

For every allocated role, each day's shift distribution must total 100%, each day's allocated productive minutes must equal that day's `RoleDailyTargetMinutes`, weekly allocated productive minutes must equal `RoleWeeklyTargetMinutes`, and reconstructed fortnight minutes must equal `RoleTargetMinutes`, within a tolerance of 0.000001.

### `MinuteWorkersFTE_CATEGORY_DAILY_CHECK`

The category proof table. Each weekday must have `AllocatedDayProductiveMinutes = CategoryWeekdayTargetMinutes`, its weighted target. The seven-day mean must equal `CategoryDailyTargetMinutes`; doubling the weekly sum must equal the fortnight target. `DayVsDailyTarget%` retains its existing name but compares with the daily average, so it need not equal 100% on each day.

### `MinuteWorkersFTE_HISTORICAL_DAY_CHECK`

The historical comparison table places average distinct workers, roster-hour FTE, and target FTE at the same facility/role/weekday grain. Historical distinct workers are headcount context, not a staffing minimum. `RoleDayTargetStatus`, direct-care reconstruction, roster-FTE arithmetic, missing employee codes, and incomplete historical coverage are shown explicitly.

### `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE`

Unaveraged graph-ready history at `Facility × Role × Week No × Weekday × Shift`. It exposes historical roster hours, roster FTE, productive hours, source-row counts, invalid-hour counts and coverage/cell flags. With two observed weeks, each observed facility/role has 42 cells (14 days × three shifts). Configured roles with no history anywhere are not invented.

`OBSERVED` cells retain source hours; `ZERO` cells are absent role/shift combinations in a complete facility week. `MISSING` cells in incomplete weeks and `ERROR` cells containing invalid raw hours have null historical FTE. Coverage completeness is inferred from eligible source weekdays, not proof that an upstream extract contains every employee. Review flags and `MW Distribution Check` before graphing. The diagnostic history output remains available independently of target validity.

## Last observed target values

The original input values below were recorded on the snapshot date. Their units were corrected on 7 September; these are calculated expectations, not results of a refreshed workbook. Reconfirm the raw inputs before using them as regression fixtures.

| Facility | RN input hours/fortnight | ALL input hours/fortnight | OTHERS hours/fortnight | RN minutes/fortnight | ALL minutes/fortnight | OTHERS minutes/fortnight |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| BD | 1041.117 | 5180.46666666667 | 4139.34966666667 | 62467.02 | 310828 | 248360.98 |
| JE | 395 | 1954 | 1559 | 23700 | 117240 | 93540 |
| TE | 738 | 3641 | 2903 | 44280 | 218460 | 174180 |

For BD RN at 100% direct care, the expected average daily productive minutes are 4461.93. The seven-day average of the AM + PM + NS FTE totals is approximately 9.784934211; individual weekdays may be higher or lower. The representative week's shift-equivalent sum is approximately 68.494539474. Display rounding only; calculations retain full precision. These are arithmetic expectations from the recorded input, not refreshed results of this revision.

The facility code is `TE`; an older `NR`/`TE` mismatch was reported as resolved.

## Last observed historical role coverage

Counts below are diagnostic row counts after role normalization, not target amounts.

| Facility | AIN | CCCM | CSM | EN | RN | RSM | WCSO |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| BD | 446 | 10 | 10 | 112 | 114 | 10 | 30 |
| JE | 154 | 9 | 10 | 28 | 40 | 10 | 14 |
| TE | 323 | 7 | 20 | 25 | 83 | 20 | 43 |

## Earlier scale correction (spot checks superseded)

Earlier work assumed the inputs were daily minutes. That assumption was incorrect and is superseded by the user's 7 September confirmation of hours per fortnight. The correct conversion is input hours times 60 for fortnight minutes, divided by 14 for daily minutes.

Examples after the scale correction, but before the weekday-local correction, were approximately:

| Facility | Role | Weekday/shift | FTE |
| --- | --- | --- | ---: |
| BD | AIN | Monday AM | 2.939611794 |
| BD | RN | Monday AM | 1.127048720 |
| JE | RN | Monday AM | 0.293941662 |
| TE | RN | Monday AM | 0.239175524 |

These values must not be used as regression expectations for the revised allocator. They predate the corrected input units and/or denominator safeguards. Refresh the workbook and establish new spot checks only after the whole-period and raw-hours reconciliations pass.

## Historical validation record (before the 7 September correction)

- Power Query query/header and dependency presentation was reviewed.
- Shared-query names were checked for duplicates.
- Delimiter counts were checked for balance.
- `git diff --check` passed apart from the repository's CRLF warning.
- The shared Power Query source validator was run with a temporary project configuration and passed six `.m` files.
- The temporary validator configuration was removed.

The shared validator is structural and does not prove Power Query runtime correctness. A language-server check was attempted earlier but timed out, so do **not** record that as a pass. Excel refresh and reconciliation remain required.

The session's earlier claim that `refresh-residentialcare.ps1 -ValidateOnly` proved this sidecar structurally valid was not supported. That command checks refresh configuration, and the default `pq.project.json` source roots cover `Workflows`, not this sidecar. Validation of the correction must explicitly include this file; source checks and independent arithmetic tests are distinct from Excel runtime execution.

## Whole-period revision source validation (7 September)

- `scripts/test-minuteworker-period-allocation.mjs`: 134 independent arithmetic/source-contract assertions passed, including unequal weekday weights, pooled-fortnight equivalence, two-stage allocation, raw-hours reconstruction, historical zero cells, invalid/missing history, and retaining more than two source weeks.
- The source has 42 uniquely named shared queries, including the new unaveraged history stage/output and weighted category-weekday targets.
- The shared validator passed six `.m` files using a temporary configuration explicitly including this sidecar's directory; the temporary configuration was removed. This validator checks file presence/content and conflict markers, not M evaluation.
- These tests do not execute Power Query, prove Excel runtime correctness, or establish that BD's two historical weeks are similar. Synchronization succeeded through the approved extension; refresh and source-data comparison remain separate acceptance steps.

## Errors already corrected

- `MinuteWorkersFTE_CHECK` / `PreAllocationChecks` previously raised: `The columns of the specified table type must be nullable.` The explicit `Table.FromRecords` schemas were changed to nullable fields.
- A previous allocation failure reported only `[Table]`. Fatal output errors now include check/facility/category/role context.
- Historical role names did not fully match `INPUT MinuteWorkers`. Role normalization was expanded and applied consistently to both sides of the join.
- Target-period scaling now converts input hours per fortnight to minutes (`* 60`) and then daily minutes (`/ 14`), superseding the former daily-minute assumption.
- MinuteWorker role and shift distributions no longer use the legacy week-wide `LocRoleTOTAL` / `LocRoleDayShift%` denominator.
- Historical averages now divide summed role/day/shift hours by the facility historical-period count, so a role or whole weekday absent in one period contributes zero instead of shrinking the averaging denominator.
- Every target facility/week must contain eligible MinuteWorker history for all seven weekdays; incomplete periods are a fatal pre-allocation error.
- Category/day and role/day reconciliation errors now block `MinuteWorkersFTE_TABLE` and `MinuteWorkersFTE_MATRIX`.
- Original RN/ALL fortnight-hour reconciliation also blocks TABLE and MATRIX through `MW Publication Check`.
- Target reshaping preserves null cells, and input checks reject an empty target table or missing RN/ALL target values.

## Continue on the other PC

1. Transfer or pull the repository and verify that both the `.m` file and this handoff arrived. Do not assume a Git pull includes local or uncommitted files; run `git status --short` on both PCs or use the approved secure transfer method.
2. Read the current repository `AGENTS.md` before acting. Instructions may have changed since this snapshot.
3. Confirm the exact workbook and `.m` paths listed above. Do not silently substitute a backup, copy, same-stem workbook, or older extracted source.
4. Review the current `.m` diff and query list before importing it.
5. Check the source used by `IMPORT Master`. It contains an external workbook path that may not exist on the new PC. Update it only through the approved path/source process; do not copy the old PC's user-specific absolute path into documentation or code examples.
6. Confirm the Excel input tables still contain the expected columns and that target facility columns still include `TE`.
7. Use the approved IMPORT Extract/synchronization method. Preserve staging queries as connection-only where supported.
8. Refresh validation outputs first:
   - `MinuteWorkersFTE_CHECK`
   - `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK`
   - `MinuteWorkersFTE_CATEGORY_DAILY_CHECK`
   - `MinuteWorkersFTE_HISTORICAL_DAY_CHECK`
9. Investigate every error and any unexpected warning. Do not publish the matrix merely because it refreshes if reconciliation fails.
10. For each allocated role/day, verify `Status = PASS`, `DailyShiftDistributionTotal% = 1`, and daily/weekly/fortnight variance is effectively zero.
11. At facility/category level, confirm:
    - every weekday's RN and OTHERS allocation equals its weighted `CategoryWeekdayTargetMinutes`;
    - their seven-day averages equal the RN and `ALL - RN` average daily targets;
    - `sum(FTE * 456 * Direct Care % * 2)` reconstructs the 14-day target when summed over the representative week.
    - `sum(FTE * 7.6 * Direct Care % * 2)` reconstructs the original RN/ALL input hours, with the independent `MW Fortnight Hours Check` showing Pass.
    - BD RN's seven daily AM + PM + NS totals average approximately 9.784934211 if the input is still 1041.117 hours/fortnight and RN direct care is 100%.
    - historical graph keys are unique, source hours/row counts reconcile, and all facility/weeks have complete seven-day coverage.
12. Only then refresh/use `MinuteWorkersFTE_TABLE` and `MinuteWorkersFTE_MATRIX`.

## Suggested verification chart

First compare the actual roster weeks using `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE`:

- Axis: weekday, ordered by `DayOfWeek`.
- Values: sum of `HistoricalRosterFTE`.
- Series: original `Week No` (two series when the source is a fortnight).
- Filters: `Facility = BD`, `Role = RN`; optionally `Shift`, or sum all three shifts to compare whole days.
- Review coverage/cell flags first. Similar lines support weekday averaging; material differences warrant retaining distinct week patterns. Similarity has not yet been confirmed numerically.

Build a PivotChart from `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK`:

- Axis: weekday, sorted by the numeric day-of-week field
- Clustered columns: `DayFTEShiftTotal`
- Line on secondary axis: `CumulativeWeekProductiveMinutes` or `CumulativeWeeklyTarget%`
- Slicers: `Facility` and `Role`
- Optional slicer: `MinuteCategory`

The final cumulative percentage should reach 100% for an allocated role, while the daily columns show how the FTE requirement varies through the week.

## Safeguards and open points

- This revision adds raw null/negative-hour validation before grouping and historical graph total/uniqueness checks. Remaining review findings: shift boundaries depend on input row order, and legacy historical outputs share the configured-role preparation filter. Do not describe the legacy branch as independent of configuration changes.

- Do not inspect, unzip, directly edit, or refresh a workbook unless the user explicitly authorizes that action and the approved workflow is followed.
- Preserve all existing non-MinuteWorker queries and outputs.
- Do not recategorize RN-qualified manager roles as the `RN` MinuteCategory unless the business rule is explicitly changed. The current rule uses the canonical role name `RN` only.
- Do not give CAM or RGM invented historical shares. Resolve missing history with a business-approved method if they are expected to receive an allocation.
- Do not use earlier Monday-only figures as a regression target. Corrected units, complete-period averaging and fortnight reconciliation are all required. Whole-period weighting does not impose a historical headcount minimum such as 14 AIN; such a floor requires a separate approved business rule.
- `Week No` is the only available historical period key. It must be non-null and uniquely identify each roster period; use a composite year/week key if the source later spans repeated week numbers.
- Confirm whether the current `.m` sidecar should be promoted to a canonical file under `Workflows/`; do not make that move implicitly.
- If the workbook inputs or historical source have changed after 30 August 2026, regenerate the diagnostic values rather than treating this handoff's numbers as authoritative.
