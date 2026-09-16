# Demand Extraction: distributed FTE integration

Current Unit1 standard-FTE implementation: [settings source, formulas and validation](2026-09-16-Whiddon-Unit1-Settings-Based-FTE-Conversions.md). Source only; workbook synchronization and refresh remain pending.

Implemented in M source on 9 September 2026, with a subsequent third-role compatibility change. **The original RN/AIN implementation passed saved-table reconciliation. The EN/AINC4 follow-up is source-only; no synchronization or Excel refresh has been performed for it.**

## Current scope

The user restored the AINC4 folder and approved using the existing `.m` files without re-extraction. The priority is RN and AIN results. Enrolled-nurse rows now use **AINC4** as their role value so the existing third-role workflow receives data. RN and AIN demand amounts and formulas are unchanged. AIN4 and unrelated roles remain excluded; their allocations are not redistributed to RN or AIN.

The modeled facility is **BD / Beaudesert**, matching the existing Allocation Extraction worksheet selection. This is an implementation assumption based on the current capacity population. `Demand Facility` makes the selection explicit. JE and TE are excluded, rather than combined with Beaudesert capacity.

The roster horizon remains **28 days**. The distinct 14-day distribution repeats twice: Week 1, Week 2, Week 1, Week 2. Each fortnight keeps the published amounts.

Changed sources (relative to the repository root):

- `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/1-AllocationExtracted.xlsx_PowerQuery.m`: emit AINC4 for Enrolled Nurse allocations so the existing role filter retains them.
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/2-DemandExtract.xlsx_PowerQuery.m`: retain the third role and replace its EN/Enrolled Nurse output value with AINC4; expand completeness checks to three roles.
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m`: replace exact EN values with AINC4 in the imported availability Role column before downstream joins.

The Manual Read allocator and its publications were not edited. The consuming Demand Extraction adapter preserves the upstream distribution. This follow-up does not edit role workbooks, runner profiles, Tableau files or any Excel workbook.

## Input and output contracts

Paths below are relative to `CLIENT/DATExx-Whiddon/UNITS/Unit1/`.

| Input | Named tables used | Purpose |
| --- | --- | --- |
| `1. Input/Demand-MasterRoster Manual Read.xlsx` | `MinuteWorkersFTE_TABLE`, `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE`, `MinuteWorkersFTE_CHECK` | Published allocation, complete zero-coverage evidence and saved upstream checks |
| `2. Calculations/Settings Data.xlsx` | `PermutationDimensions`, `ShiftPeriod` | Calendar, existing period IDs and role-specific shift timing |

The existing settings helper queries remain. The exact settings workbook is now present; no alternative path is used.

Path setup follow-up: the saved Demand Extract workbook has a single-cell name `FilePathUrl` pointing to `Unit1Path!B1`, and a separate query-output table `UnitL1PathTABLE` at `Unit1Path!A3:B10`. The resolver now accepts this existing named-cell input as well as the standard `FilePath` table column, normalizes the `folder/[workbook]sheet` filename form, and rejects paths identifying another workbook. `UnitL1PathTABLE` remains the output destination. The saved B1 value identifies an older donor; use `=CELL("filename",A1)` in B1 and save/recalculate the current workbook before refreshing. This source edit did not change that workbook formula or synchronize M.

The Manual Read workbook is read through one shared import with a buffered binary and navigation table. The adapter supports the two known publication schemas:

- Older saved publication: `QFR Category`, with RN/AIN role abbreviations.
- Revised publication: `DC Role` and `DC Category`, with explicit full role labels.

Allocation and profile must have compatible schema markers and required columns. Selected labels are RN / Registered Nurse, AIN / Assistant in Nursing, and EN / Enrolled Nurse / AINC4, normalized for case and surrounding whitespace. Enrolled-nurse demand is published under AINC4. No role is inferred from category alone. This compatibility does not assert that saved tables reflect the latest source edit.

The old `IMPORT LocRoleDayShift%` and `LocRoleDayShift Prepare` queries are replaced. The existing `Permutation DateTimeRoleShift` query name is retained with the new timing preparation.

The published query remains **`ShiftUnitDemandHRS`**, with these exact columns in order:

```text
Date, Day, Shift, Period, Role, StartTime, EndTime, Unit, Facility,
DemandFTE, DemandHRS, DurationOfShifts
```

There are **252 rows**: 28 dates × three roles × three shifts, sharing **84 existing period IDs**. The RN/AIN subset remains 168 rows. Both Facility and Unit identify BD. Validated zero-demand rows retain their timing.

When synchronization is explicitly authorized later, retain the existing output worksheet/table name and headers at row 1: downstream consumers import the sheet. New staging queries should be connection-only; diagnostic outputs can be loaded separately.

## Calculation

```text
FortnightDayIndex = ((Date - planning start) modulo 14) + 1
ShiftDuration = Settings ShiftDuration hours
DemandHRS = published source FTE × ShiftDuration
DemandFTE = DemandHRS / DurationOfShifts
```

`Day` remains the settings sequential index 1–28. It is not used as a weekday join key. The source has 14 distinct days; planning days 15–28 map back to source days 1–14.

The settings calendar must contain exactly 28 consecutive dates, begin on Monday and provide periods 1–84 shared by RN, AIN and AINC4. The adapter neither pads nor truncates an invalid calendar. A different starting weekday needs an explicit pattern-alignment change.

Source FTE represents the configured standard roster hours. For example, with a 7.6-hour setting and an 8.25-hour shift, one source FTE becomes approximately 0.921212 shift-average attendance. Full precision is retained. Direct Care % is applied only in productive-hour reconciliation, not a second time to roster demand.

Night shifts end on the following day. Every settings duration must equal its end timestamp minus its start timestamp, including zero-demand cells.

## Validation and audit queries

| Query | Check or output |
| --- | --- |
| `Distributed FTE Source Validate` | Required columns, matching schema family, ShiftDuration in hours matching current settings, and nonempty saved upstream checks with no failures |
| `Distributed FTE Rows Prepare` | Source keys, finite values, fortnight alignment, direct-care percentages and duplicate keys |
| `Distributed FTE Prepare` | Exactly 126 unique RN/AIN/AINC4 fortnight cells for BD, two distinct ordered source weeks, allocation/profile agreement and proven zeros |
| `Demand Calendar Prepare` | Exactly 28 dates, 252 role/shift cells and 84 consistent periods |
| `Permutation DateTimeRoleShift` | Exactly one timing match per calendar cell and correct durations |
| `Demand Extraction Prepare` | Exactly one pattern match per calendar cell and reversible hours/attendance conversion |
| `DemandExtraction_RECONCILIATION` | Roster, productive and integrated hours for each role in each fortnight |
| `Distributed FTE Exclusions` | Excluded source allocations and their fortnight roster/productive hours, by facility and original role |
| `DemandExtraction_CHECK` | Readable stage failures and upstream warnings; any error blocks `ShiftUnitDemandHRS` |

A missing sparse allocation becomes zero only when the complete profile proves a PASS ZERO cell, zero historical/redistributed FTE and no allocation match. Missing positive allocations, duplicated matches and mismatched original weeks remain errors.

Each retained role/cycle must contain 42 cells and reconcile to its published fortnight pattern, giving six reconciliation rows. The 28-day total is twice that pattern. Retained roles are not required to consume the entire upstream ALL target, because unrelated allocated roles remain excluded.

## Original RN/AIN validation performed

- Installed Microsoft Power Query language service: no diagnostics after loading the repository's Excel function symbols. An intentionally invalid control query produced a parse error, verifying the validator was active.
- `scripts/test-demand-extraction-distributed-fte.py`: five test groups passed, covering unequal weeks, repetition, fractional attendance, overnight timing, zero cells, excluded roles/facilities, both role-label formats, malformed publications and invalid calendars/durations.
- With `--saved-workbooks`, the same script independently read the exact named source/settings tables and reconstructed the expected 168-row result. It checked the saved calendar, period IDs and shift spans.
- Existing MinuteWorker suite: 544 arithmetic/source-contract assertions passed.
- Diff whitespace checks passed. New query headers, final `in` targets, dependency flow, public columns and publication gating were reviewed.

Saved-table reference results for BD:

| Role | Roster hours per fortnight | Roster hours over 28 days |
| --- | ---: | ---: |
| RN | 1041.117000000 | 2082.234000000 |
| AIN | 3049.510542781 | 6099.021085562 |

These are independent calculations from saved inputs, not refreshed Demand Extraction outputs. The Python fixtures do not execute M. Syntax validation does not prove Excel runtime behavior.

## EN/AINC4 follow-up validation

- All three edited M sources passed Microsoft's Power Query parser. Existing shared query names and public output interfaces were preserved.
- Independent JavaScript fixtures produced 252 demand rows over 84 periods. Filtering the result to RN/AIN exactly reproduced the previous 168 rows, including fractional demand and fortnight repetition. Missing or duplicate third-role coverage was rejected by the reference checks.
- `scripts/test-demand-extraction-distributed-fte.py` was updated for three roles, including EN/Enrolled Nurse/AINC4 labels, RN/AIN preservation and missing/duplicate third-role cases. It was not executed in this session because a Python interpreter was unavailable. The JavaScript checks do not execute M.
- Diff whitespace checks passed. Query headers, comments, dependencies, final `in` targets and validation gates were reviewed.
- No Excel workbook was inspected, edited, synchronized or refreshed for this follow-up. The earlier saved-table results above describe the original two-role version only.

## Remaining workbook validation

Source implementation is complete. Synchronization and refresh remain separate, explicitly authorized operations under AGENTS.md.

After authorized synchronization, preserve the main worksheet/table interface and verify `DemandExtraction_CHECK` and `DemandExtraction_RECONCILIATION` in Excel. Re-extract the synchronized workbook through the approved mechanism and compare with this source. Then, when refresh is authorized, follow the approved dependency order and confirm fractional demand reaches the RN and AIN A.1 models.

The follow-up includes AINC4 demand using the enrolled-nurse role value, fixes the Allocation Extraction EN-filter issue by emitting AINC4, and replaces EN in imported availability. Complete third-role source coverage and AINC4 settings/timing are still required. Missing third-role inputs fail visibly rather than being replaced with invented values. The restored folder and M syntax have been checked; workbook contents and refresh behavior have not been verified for this follow-up.

Known unchanged downstream limitations:

- `Demand.xlsx` casts its diagnostic `DemandFTE` column to an integer. The reviewed A.1 path uses the separate numeric `EffectiveIntervalAttendance` measure; verify that path at runtime.
- Three legacy demand checks divide whole-horizon hours by two while labeling the result weekly. Over 28 days those values represent a fortnight average. Use the new explicit fortnight reconciliations for acceptance.
- `MaxAvailability = 10` and `StdRosterDays = 10` are capacity controls, not date-range generators. They were not changed by this task.

## Analysis-period findings retained from review

Settings derives its start from the first allocation date on or after `DateFrom` and its end from the maximum eligible allocation date. No separate fixed analysis-duration parameter was found in the reviewed settings source. The adapter now enforces the user-confirmed 28-day duration.

The saved settings calendar is 20 July–16 August 2026. The earlier inspected settings candidate had `DateFrom = 15 July 2024`, which precedes that roster and does not shorten it. The adapter anchors repetition to the actual first planning date, not that raw cutoff.

Manual Read's `TargetPeriodDays = 14` remains unchanged. Ordinary-hour settings of 38/week, 76/fortnight and 152/28 days describe workload quantities, not calendar generation.
