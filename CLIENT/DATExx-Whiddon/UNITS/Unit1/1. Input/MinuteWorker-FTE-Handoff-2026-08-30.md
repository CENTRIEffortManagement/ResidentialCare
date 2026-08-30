# MinuteWorker FTE Analysis — PC Handoff

**Snapshot date:** 30 August 2026, Australia/Sydney.

> **Staleness warning:** This document describes the work as it stood on the date above. If Codex or another developer reads it much later, it may be out of date. Before making changes, re-read the repository `AGENTS.md`, check the current Git/worktree state, inspect the current `.m` source, and verify the workbook table structures and input values.

## Goal

Produce a MinuteWorker-only analysis that allocates each facility's target productive-care minutes to roles, weekdays, and shifts using historical roster patterns, then converts the required roster minutes to FTE.

The analysis is a separate branch. Existing roster-analysis outputs should remain unchanged.

## Current working files

- Power Query source: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m`
- Target workbook: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx`
- This handoff: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/MinuteWorker-FTE-Handoff-2026-08-30.md`

During this work, the user explicitly instructed Codex to use the current `.m` file and the approved **IMPORT Extract** method. Do not overwrite the current `.m` by extracting older code from the workbook. Confirm the source-of-truth state before any future extraction or synchronization. Repository policy normally places canonical product-specific Power Query source under `Workflows/`, so reconcile that policy with the explicitly selected sidecar source before moving or syncing anything.

Codex did **not** write to or refresh the Excel workbook. The changes described here are in the `.m` source. Import/synchronization and workbook refresh still need to be completed through the approved process.

## Confirmed business rules

1. `INPUT MinuteWorkers` is the role list and contains each role's `Direct Care %`.
2. `INPUT TargetMinutes` contains **daily** target minutes, not fortnight totals.
3. Targets are supplied for each facility for two types:
   - `RN`
   - `ALL`, which includes RN minutes
4. The calculated MinuteCategories are:
   - `RN`: the canonical role `RN` only
   - `OTHERS`: every other MinuteWorker role, including RN-qualified management roles
5. Daily category targets are:
   - `RNDailyTargetMinutes = RN`
   - `OTHERSDailyTargetMinutes = ALL - RN`
6. Fortnight category targets equal the daily target multiplied by 14.
7. Historical hours are filtered to MinuteWorker roles and adjusted for direct care:
   - `HistoricalMinuteHours = HistoricalRosterHours * Direct Care %`
8. Each role's share inside its facility and MinuteCategory is:
   - `RoleHistoryDistribution% = Role HistoricalMinuteHours / Category HistoricalMinuteHours`
9. Role target minutes are allocated through weekdays and the three shifts using the existing historical `LocRoleDayShift%`.
10. The calculation builds one representative week. Because the role target is a 14-day target:
    - `WeekdayShiftTargetMinutes = RoleTargetMinutes * LocRoleDayShift% / 2`
11. Productive-care target minutes are converted back to roster minutes:
    - `WeekdayShiftRosterMinutes = WeekdayShiftTargetMinutes / Direct Care %`
12. One FTE shift is 7.6 hours or 456 minutes:
    - `FTE = WeekdayShiftRosterMinutes / 456`
13. Calculations retain full precision. Rounding is for display only.
14. A positive facility/category target with no eligible historical hours must fail validation. The calculation must not invent a distribution.

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
- existing historical weekday/shift queries, ending in `LocRoleDayShift%`
- `INPUT TargetMinutes`
- `MW TargetMinutes Prepare`

### Historical distribution

- `MW Historical DayShift`
- `MW Role Distribution`

### Allocation

- `MW Role Targets`
- `MW DayShift Allocation`

### Validation

- `MW Input Check`
- `MW Distribution Check`
- `MW PreAllocation Check`

### Published outputs

- `MinuteWorkersFTE_TABLE`
- `MinuteWorkersFTE_MATRIX`
- `MinuteWorkersFTE_CHECK`
- `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK`

Staging queries should be connection-only in Excel when supported. Load only the outputs that users need to worksheets or the data model.

## Output grains and purpose

### `MinuteWorkersFTE_TABLE`

Detailed allocation at approximately:

`Facility × MinuteCategory × Role × Weekday × Shift`

It contains the historical measures and distributions, direct-care percentage, daily/fortnight/role targets, allocated productive minutes, roster minutes, and unrounded FTE.

### `MinuteWorkersFTE_MATRIX`

One row per facility/category/role, with weekday-shift combinations as FTE columns. This is the user-facing staffing pattern view.

### `MinuteWorkersFTE_CHECK`

The principal validation output. It surfaces input, matching, distribution, allocation, and reconciliation failures. Required failures must stop the detailed output from silently publishing invalid results.

### `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK`

The proof table requested most recently. It has seven rows for every facility/category/role, including zero-allocation days, and shows that variable daily distributions still reconcile to the weekly and fortnight targets.

Important columns include:

- `DayDistribution%`
- `AllocatedDayProductiveMinutes`
- `AllocatedDayRosterMinutes`
- `DayFTEShiftTotal`
- `ExpectedWeeklyProductiveMinutes`
- `AllocatedWeeklyProductiveMinutes`
- `WeeklyVarianceMinutes`
- `WeeklyDistributionTotal%`
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

For a valid allocated role, the weekly distribution must total 100%, weekly allocated productive minutes must equal `RoleDailyTargetMinutes * 7`, and the reconstructed fortnight minutes must equal the 14-day `RoleTargetMinutes`, within a tolerance of 0.000001.

## Last observed target values

These values were read from the workbook on the snapshot date and are included only as continuation spot checks. Reconfirm them if the inputs have changed.

| Facility | RN daily | ALL daily | OTHERS daily | RN fortnight | ALL fortnight | OTHERS fortnight |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| BD | 1041.117 | 5180.46666666667 | 4139.34966666667 | 14575.638 | 72526.53333333338 | 57950.89533333338 |
| JE | 395 | 1954 | 1559 | 5530 | 27356 | 21826 |
| TE | 738 | 3641 | 2903 | 10332 | 50974 | 40642 |

The facility code is `TE`; an older `NR`/`TE` mismatch was reported as resolved.

## Last observed historical role coverage

Counts below are diagnostic row counts after role normalization, not target amounts.

| Facility | AIN | CCCM | CSM | EN | RN | RSM | WCSO |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| BD | 446 | 10 | 10 | 112 | 114 | 10 | 30 |
| JE | 154 | 9 | 10 | 28 | 40 | 10 | 14 |
| TE | 323 | 7 | 20 | 25 | 83 | 20 | 43 |

## Scale correction already made

An earlier version incorrectly treated `INPUT TargetMinutes` as a 14-day total, producing FTE values that were too low. The user confirmed the inputs are per day, so `TargetPeriodDays = 14` is now applied before distributing to the representative week.

Examples after the correction were approximately:

| Facility | Role | Weekday/shift | FTE |
| --- | --- | --- | ---: |
| BD | AIN | Monday AM | 2.939611794 |
| BD | RN | Monday AM | 1.127048720 |
| JE | RN | Monday AM | 0.293941662 |
| TE | RN | Monday AM | 0.239175524 |

These are regression spot checks, not fixed expected values: they may change if historical data, inputs, or mappings change.

## Validation already performed on the `.m` source

- Power Query query/header and dependency presentation was reviewed.
- Shared-query names were checked for duplicates.
- Delimiter counts were checked for balance.
- `git diff --check` passed apart from the repository's CRLF warning.
- The shared Power Query source validator was run with a temporary project configuration and passed six `.m` files.
- The temporary validator configuration was removed.

The shared validator is structural and does not prove Power Query runtime correctness. A language-server check was attempted earlier but timed out, so do **not** record that as a pass. Excel refresh and reconciliation remain required.

## Errors already corrected

- `MinuteWorkersFTE_CHECK` / `PreAllocationChecks` previously raised: `The columns of the specified table type must be nullable.` The explicit `Table.FromRecords` schemas were changed to nullable fields.
- A previous allocation failure reported only `[Table]`. Fatal output errors now include check/facility/category/role context.
- Historical role names did not fully match `INPUT MinuteWorkers`. Role normalization was expanded and applied consistently to both sides of the join.
- Target-period scaling was corrected from fortnight input to daily input multiplied by 14.

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
9. Investigate every error and any unexpected warning. Do not publish the matrix merely because it refreshes if reconciliation fails.
10. For each allocated role, verify `Status = PASS`, `WeeklyDistributionTotal% = 1`, and weekly/fortnight variance is effectively zero.
11. At facility/category level, confirm:
    - RN allocations reconcile to the RN target.
    - OTHERS allocations reconcile to `ALL - RN`.
    - `sum(FTE * 456 * Direct Care % * 2)` reconstructs the 14-day target when summed over the representative week.
12. Only then refresh/use `MinuteWorkersFTE_TABLE` and `MinuteWorkersFTE_MATRIX`.

## Suggested verification chart

Build a PivotChart from `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK`:

- Axis: weekday, sorted by the numeric day-of-week field
- Clustered columns: `DayFTEShiftTotal`
- Line on secondary axis: `CumulativeWeekProductiveMinutes` or `CumulativeWeeklyTarget%`
- Slicers: `Facility` and `Role`
- Optional slicer: `MinuteCategory`

The final cumulative percentage should reach 100% for an allocated role, while the daily columns show how the FTE requirement varies through the week.

## Safeguards and open points

- Do not inspect, unzip, directly edit, or refresh a workbook unless the user explicitly authorizes that action and the approved workflow is followed.
- Preserve all existing non-MinuteWorker queries and outputs.
- Do not recategorize RN-qualified manager roles as the `RN` MinuteCategory unless the business rule is explicitly changed. The current rule uses the canonical role name `RN` only.
- Do not give CAM or RGM invented historical shares. Resolve missing history with a business-approved method if they are expected to receive an allocation.
- Confirm whether the current `.m` sidecar should be promoted to a canonical file under `Workflows/`; do not make that move implicitly.
- If the workbook inputs or historical source have changed after 30 August 2026, regenerate the diagnostic values rather than treating this handoff's numbers as authoritative.
