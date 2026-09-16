# Shift duration, net hours and FTE reconciliation handover

**16 September 2026 standard-FTE implementation:** Unit1 Manual Read and Demand Extract now use Settings Data's ShiftDuration directly in hours. See [formulas, exact scope, publication prerequisites and validation](2026-09-16-Whiddon-Unit1-Settings-Based-FTE-Conversions.md). Earlier fixed-denominator observations below are historical. The approved net-hours numerator remains; actual role/shift duration is still Demand Extract's attendance denominator. E-O-I's Effort-All remains excluded. No workbook was inspected, synchronized or refreshed.

Date: 2026-09-15

File-impact register expanded: 2026-09-16. See [Files affected by shift hours, meals, headcount and FTE](Shift-Hours-FTE-Impacted-Files.md) for the complete identified input, calculation and reporting chain, including retained variants and evidence limits.

## Purpose

Handover for continuing this discussion on another computer. The user asked whether the demand model uses `Roster Hours` or `Shift Net Length`, whether meals are removed, and whether downstream `2-DemandExtract` incorporates meals.

Scope: `CLIENT/DATExx-Whiddon/UNITS/Unit1/` only. Findings below come from reading the saved Power Query `.m` files. The embedded workbook queries, worksheet formulas, saved data, and refresh currency have not been verified. No workbooks were inspected, changed, synchronized, or refreshed. The original 15 September review made no M-code changes. On 16 September the user approved changing Manual Read to use `Shift Net Length`; that source-only change is recorded below.

**16 September source update:** `Master Prepare` now selects `Shift Net Length` and renames it to the existing internal `Roster Hours` column. Historical distributions and checks therefore use the explicit net field. There is no fallback to the original `Roster Hours` or additional meal deduction. Shift splitting, synchronization and refresh were not part of this change. Line references in the original review predate this edit.

## Current agreed direction

- The roster's explicit net-hours value is the authority for rostered workers: rostered clock time less unpaid meal time.
- `Shift FTE = net hours allocated to the shift / configured gross shift hours`.
- Shift hours may vary by facility, role and AM/PM/NIGHT period. The intended settings key is therefore `Facility + Role + ShiftPeriod` where the configuration supports it.
- Start and end times define presence and shift-boundary overlap. They must not calculate a second meal deduction when source net hours already exist.
- Interval values should be FTE rates. Shift FTE is their duration-weighted average.
- Demand begins as a shift-average requirement and is repeated across the intervals in that shift.
- Allocation begins as individual roster rows, so each worker's source net value must first be distributed over the intervals they overlap.
- Effort-All should consume reconciled FTE without another fixed `7.6 / 8` reduction.
- Availability Redistribution remains separate because it starts with nominal availability FTE rather than observed roster net hours. For a qualifying shift, the intended policy is a 0.5-hour reduction when the shift is longer than six hours.

## Related artifacts

- [Full static reconciliation](../outputs/shift-duration-reconciliation/README.md)
- [Detailed evidence index](../outputs/shift-duration-reconciliation/index.md)
- [Shift Duration chart](Shift-Duration-Chart.mmd)
- [Original ResidentialCare file/table map](ResidentialCare-File-Table-Connections.mmd)

## Definitions to keep separate

| Name | Meaning | Unit | Intended use |
| --- | --- | --- | --- |
| `RosterNetHours` | Source roster time after unpaid meal time | hours | Authoritative worker quantity |
| `RosterClockHours` | End timestamp minus start timestamp | hours | Presence boundaries and proportional interval distribution |
| `WorkerFTERate` | `RosterNetHours / RosterClockHours` | FTE | Worker's average contribution while rostered |
| `ShiftHours` | Configured gross AM/PM/NIGHT span for the facility and role | hours | Denominator for shift-average FTE |
| `IntervalFTE` | Sum of active worker FTE rates, or Demand FTE repeated into an interval | FTE | RosterProfile and interval comparison |
| `IntervalFTEHours` | `IntervalFTE * IntervalHours` | FTE-hours | Additive internal measure used for correct aggregation |
| `ShiftFTE` | `sum(IntervalFTEHours) / ShiftHours` | FTE | Demand, Allocation, Effort and efficiency comparisons |
| `StandardFTEHours` | Separate standard-hours convention, currently 7.6 in several places | hours per standard FTE | HR/resource and cost conversions only when explicitly required |
| `ProductiveCareMinutes` | Target direct-care requirement | minutes | Demand budget before Direct Care % conversion |
| `DemandRosterMinutes` | Productive minutes divided by Direct Care % | minutes | Demand time requirement; its net/gross meaning still needs confirmation |

## Was the explicit net value used?

| Path | Current source or calculation | Was explicit net authoritative? | Revised decision |
| --- | --- | --- | --- |
| Demand historical distribution | `Shift Net Length`, renamed to internal `Roster Hours` | Yes in the M source after the approved 16 September edit; workbook parity remains unverified. | Preserve net hours; shift-boundary splitting remains separate work. |
| Demand target quantity | Target productive minutes divided by `Direct Care %` | No explicit meal calculation is applied. Net status depends on what `Direct Care %` represents. | Define the result as net roster minutes only after confirming that business meaning. |
| Demand intervals | `EffectiveIntervalAttendance = DemandFTE` | The separately calculated meal ratio is bypassed. | Keep shift FTE unchanged across its intervals; do not deduct a meal again. |
| Allocation extraction | `Shift Net Length` renamed to `Hours` | The field is selected and carried at the input. | Rename clearly to `RosterNetHours` and retain a stable source-row identity. |
| Allocation intervals | `Effective Duration / RealDuration` after Shifts recalculates meal time | No. The source net field is not the controlling numerator. | Use the source `RosterNetHours / RosterClockHours` rate. |

## 1. Master-roster demand source

Source: [Demand-MasterRoster Manual Read.xlsx_PowerQuery.m](../CLIENT/DATExx-Whiddon/UNITS/Unit1/1.%20Input/Demand-MasterRoster%20Manual%20Read.xlsx_PowerQuery.m)

- `IMPORT Master` imports `Roster Hours`, `Shift Net Length`, and break fields from the `Combined` sheet of `Master Roster.xlsx` (lines 8–11).
- `Master Prepare` now selects `Shift Net Length` and renames it to `Roster Hours`; the original source `Roster Hours` and break fields do not enter historical calculations.
- `LocRoleWeekDaysHours` sums the net-hours `Roster Hours` alias.
- `MW Historical WeekDayShift` also sums that net-hours alias.
- Historical FTE is `HistoricalRosterHours / 7.6` (lines 703–705).
- Historical productive hours are `HistoricalRosterHours * [Direct Care %]` (lines 706–707).

**Current source finding:** The duration basis is now `Shift Net Length`. The M source applies no additional meal deduction. `Direct Care %` remains a separate configured adjustment. The meaning of the original source `Roster Hours` remains unverified, but that field no longer drives historical calculations.

The published `MinuteWorkersFTE_TABLE` is a target-based allocation using the historical profile, expressed as 7.6-hour roster shift equivalents. It is not simply a copy of raw historical hours.

## 2. Conversion in 2-DemandExtract

Source: [2-DemandExtract.xlsx_PowerQuery.m](../CLIENT/DATExx-Whiddon/UNITS/Unit1/1.%20Input/2-DemandExtract.xlsx_PowerQuery.m)

The source reads saved tables from `Demand-MasterRoster Manual Read.xlsx`:

- `MinuteWorkersFTE_TABLE`
- `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE`
- `MinuteWorkersFTE_CHECK`

It validates the allocation against the paired profile and upstream checks. The numerical source for `SourceFTE` is the profile's `RedistributedRosterFTE`, checked against the allocation's `FTE`; missing allocation cells require explicit zero evidence.

Transformations:

1. Select facility `BD` and roles `RN`, `AIN`, and `AINC4`; map `EN`/`ENROLLED NURSE` to `AINC4` and shift `NS` to `NIGHT`.
2. Convert source FTE to hours: `DemandHRS = SourceFTE * 7.6` (line 381).
3. Repeat the 14-day source pattern over the 28-day settings calendar, once per fortnight.
4. Read shift timing from `Settings Data.xlsx` and require `DurationOfShifts` to equal the full `EndTime - StartTime` span (lines 435–440).
5. Calculate average attendance: `DemandFTE = DemandHRS / DurationOfShifts` (line 460).
6. Reconcile integrated attendance hours to the source demand hours for each fortnight.
7. Publish `ShiftUnitDemandHRS` for downstream use.

`Direct Care %` is not applied again to demand attendance. It is retained for validation and productive-hour reconciliation.

**Finding:** This stage changes units and distributes hours across the full shift span; it does not subtract meals.

Illustrative example, not a verified settings value: 1 source FTE becomes 7.6 demand hours. Over an 8-hour shift, average attendance is 7.6 / 8 = 0.95. Integrating 0.95 across 8 hours still gives 7.6 hours. This conversion is not an additional meal deduction.

## 3. Downstream DemandIntervals meal calculation

Source: [DemandIntervals.xlsx_PowerQuery.m](../CLIENT/DATExx-Whiddon/UNITS/Unit1/2.%20Calculations/DemandIntervals.xlsx_PowerQuery.m)

`IMPORT Table_ShiftDemandHRS` reads the `ShiftUnitDemandHRS` sheet from `2-DemandExtract.xlsx` and renames `DurationOfShifts` to `ShiftDuration` (lines 92–101).

In `ShiftDemandUnitINTERVAL` (lines 163–171):

```powerquery
// Existing logic, abbreviated to show the formulas.
EffectiveDuration = if [ShiftDuration] < MealBreakStart
    then [ShiftDuration]
    else [ShiftDuration] - 0.5
IntervalEffectiveRatio = [EffectiveDuration] / [ShiftDuration]
```

The multiplication that would apply this meal ratio is commented out:

```powerquery
// each [DemandFTE] * [IntervalEffectiveRatio]
```

The active calculation uses `DemandFTE` unchanged:

```powerquery
// Active formulas, abbreviated.
EffectiveIntervalAttendance = [DemandFTE]
DemandEffort = [Duration] * [EffectiveIntervalAttendance] * 24
```

The adjacent existing comment says the meal adjustment is not applicable for ANACC demand and is only for master-roster demand. This documents the existing rationale, not a newly approved business rule.

**Finding:** A meal-adjusted duration and ratio are calculated, but that ratio is not applied to demand attendance or demand effort. The deduction is hard-coded as 0.5 hours at this step, although a `MealBreakTime` query also exists.

An additional check query, `Table_ShiftDemandHRSCheck`, calculates an approximate care-hours figure using a factor of `0.936` (line 208). This is a separate check output; it does not adjust the demand effort calculation traced above.

## Static conclusion and revised decision

The revised source uses `Shift Net Length` as its historical duration basis under the internal `Roster Hours` alias, produces target-based roster FTE, converts that FTE to demand hours and average attendance, and does not apply an additional meal deduction to the resulting demand effort. Saved workbook publications have not been updated by this source edit.

The prior source used raw `Roster Hours` without an explicit meal correction. Switching to net history can change the allocation weights. Because final FTE is target-scaled, assess the effect through both the historical weights and the target/direct-care definitions rather than assuming a one-for-one reduction of final demand hours.

The subsequent discussion established the intended direction: use the roster's explicit net value as the allocation authority, while retaining start/end timestamps only for temporal distribution. The remaining questions are about field semantics, configuration keys and downstream consumers, not whether roster meal time should be deducted again.

Before implementation, establish:

- Whether `Roster Hours` is identical to the authoritative `Shift Net Length`, gross, or differently defined, including units.
- What the approved `Direct Care %` values are intended to cover.
- Whether demand should represent roster attendance, working time excluding meals, or productive direct-care time at each interface.
- Whether the current embedded workbook queries match these saved M files.

Do not enable the downstream meal ratio. Once the approved source net field is used, doing so would deduct meals again.

Follow repository `AGENTS.md` for any continuation. Workbook inspection requires explicit authorization; workbook-linked edits require the approved source-of-truth process. Synchronization and refresh remain separate opt-in operations.

Line numbers refer to the files as read during this discussion and may move after later edits.

## Why the Demand 7.6 conversion can be simplified

The current Demand transformation is:

```text
SourceFTE = DemandRosterMinutes / 456
DemandHRS = SourceFTE * 7.6
DemandFTE = DemandHRS / actual ShiftHours
```

Because `456 = 7.6 * 60`:

```text
DemandFTE
    = DemandRosterMinutes / (60 * actual ShiftHours)
```

The intermediate `/456` and `*7.6` cancel. The eventual `DemandFTE` is already average attendance across the actual shift. A clearer target is:

```text
DemandShiftFTE = DemandRosterMinutes / configured ShiftMinutes
```

The 7.6-hour equivalent can remain as a separately named compatibility or HR measure if a consumer requires it. It should not be confused with actual-shift attendance FTE.

## When a worker is distributed into AM, PM and NIGHT

There are currently two operations:

1. `StartTime+Duration` in [Intervals.xlsx Power Query](../CLIENT/DATExx-Whiddon/UNITS/Unit1/2.%20Calculations/Intervals.xlsx_PowerQuery.m) gives the whole roster row one coarse label from its start time. This does not split time.
2. `Intervals` constructs interval boundaries and labels each interval AM, PM or NIGHT. `NAMES+INTERVALS` in [Shifts.xlsx Power Query](../CLIENT/DATExx-Whiddon/UNITS/Unit1/2.%20Calculations/Shifts.xlsx_PowerQuery.m) expands a worker from their start interval to their end interval and inherits those labels. This is the meaningful distribution.

Demand's historical master-roster path currently performs only the first operation. It assigns the whole row by start time. The revised design should segment master-roster history with the same boundary method used for Allocation before historical Demand weights are calculated.

## Revised interval calculations

### Allocation

Calculate the FTE rate once at individual roster-row grain:

```text
WorkerFTERate = RosterNetHours / RosterClockHours
```

Split the row wherever it intersects an interval or configured shift boundary. Each overlapped interval inherits the worker's rate:

```text
WorkerIntervalFTE = WorkerFTERate
AllocationIntervalFTE = sum(active WorkerIntervalFTE)
```

The additive internal measure is:

```text
AllocationIntervalFTEHours
    = AllocationIntervalFTE * IntervalHours
```

Aggregate to shift as a duration-weighted average:

```text
AllocationShiftFTE
    = sum(AllocationIntervalFTEHours) / configured ShiftHours
```

This equals:

```text
AllocationShiftFTE
    = net roster hours allocated to the shift / configured ShiftHours
```

No other meal factor should be applied.

### Demand

After target minutes have been allocated to the historical net-roster profile:

```text
DemandShiftFTE = DemandRosterMinutes / configured ShiftMinutes
DemandIntervalFTE = DemandShiftFTE
```

Demand is an average staffing requirement for the whole shift, so its shift FTE is repeated unchanged across every interval in that shift.

The interval re-aggregation can remain:

```text
ReaggregatedDemandShiftFTE
    = sum(DemandIntervalFTE * IntervalHours) / ShiftHours
```

This is not a second meal or FTE conversion. It is a weighted average. When the interval value is constant, it reproduces the original `DemandShiftFTE`. Preserve the source shift FTE and use the reaggregated result as a reconciliation unless interval-level Demand is deliberately varied.

### Crossing-shift example

Assume a worker is rostered 14:00-22:00, has 7.5 net hours, and the AM/PM boundary is 15:00:

```text
WorkerFTERate = 7.5 / 8 = 0.9375
```

| Shift segment | Clock overlap | Worker interval FTE | Allocated net hours |
| --- | ---: | ---: | ---: |
| AM | 1 hour | 0.9375 | 0.9375 |
| PM | 7 hours | 0.9375 | 6.5625 |
| Total | 8 hours | - | 7.5 |

Each segment is subsequently assessed against that facility/role/period's configured gross shift hours. If an actual break timestamp is not used, proportional spreading exactly preserves the source net total without inventing a break location.

## Reference values for an eight-hour shift

Assume an eight-hour gross shift, 0.5-hour unpaid meal and a separate 7.6-hour standard-FTE convention:

| Quantity | Value | Interpretation |
| --- | ---: | --- |
| Gross shift hours | 8.0 | Configured denominator |
| Source net roster hours | 7.5 | Authoritative worker numerator |
| Net attendance rate | 0.9375 | `7.5 / 8` |
| Actual-shift FTE | 0.9375 | Average attendance over the eight-hour shift |
| Standard resource FTE | 0.9868421053 | `7.5 / 7.6`; a different measure |
| Existing Effort-All factor | 0.95 | `7.6 / 8`; not the 0.5-hour meal result |
| One 7.6-hour Demand equivalent over eight hours | 0.95 | Valid only when the Demand source quantity is 7.6 hours |

The 0.9375, 0.9868421053 and 0.95 values answer different questions and require distinct names.

## Downstream paths

| Consumer/path | Required measure | Revised treatment |
| --- | --- | --- |
| RosterProfile interval height | `AllocationIntervalFTE` | Plot the sum of source-net worker FTE rates. Height excludes unpaid meal time without locating a specific break interval. |
| Effort-All efficiency via `EffortAllMatrixAG1_1D` | Actual role/facility Demand, Allocation and relevant capacity FTE | Consume reconciled FTE directly. Remove the additional `7.6 / 8` reduction once its existing business purpose is confirmed. |
| Availability Redistribution via `RoleAvailabilityDevelopedMATRIX` and `RoleAvailabilityDevelopedMATRIXDELTA` | Net availability FTE | Because availability often begins as nominal `1.0` with no observed net value, derive `NominalFTE * (ShiftHours - 0.5) / ShiftHours` when the shift is longer than six hours. |
| Cost | Explicitly named FTE basis | Confirm whether Cost consumes standard resource FTE or actual-shift attendance FTE before changing the hardwired 7.6 conversion. |

Roster and Demand must not apply the Availability policy when their source quantity is already net or has been defined as net.

## File-by-file cleanup map

This table summarises the principal calculation decisions. The [expanded file register](Shift-Hours-FTE-Impacted-Files.md) also names source inputs, all three capacity stages for RN/AIN/AINC4, staff/headcount checks, organisation consumers and Tableau reports. A listed downstream consumer needs validation; it does not necessarily need a formula edit.

| File | Current duration/FTE role | Core decision |
| --- | --- | --- |
| `1-AllocationExtracted.xlsx` | Imports `Shift Net Length` as `Hours` plus start, end and break length | Preserve it as `RosterNetHours`; retain a stable roster-row key. |
| `Demand-MasterRoster Manual Read.xlsx` | Source now uses `Shift Net Length` through the `Roster Hours` alias; assigns whole rows by start time and publishes a 7.6-hour equivalent | Net-field selection completed in source on 16 September. Splitting crossing rows and workbook validation remain outstanding. |
| `2-DemandExtract.xlsx` | Converts `/456 -> *7.6 -> /actual shift hours` | Calculate actual-shift Demand FTE directly from Demand roster minutes and configured shift minutes; name compatibility measures separately. |
| `Settings Data.xlsx` | Supplies boundaries, role shift duration, Meals and standard duration | Establish one authoritative `Facility + Role + ShiftPeriod` gross duration/boundary table and keep meal policy separate. |
| `Intervals.xlsx` | Creates roster and shift boundaries; also contains a duration/meal calculation with inconsistent day/hour units | Use it for temporal segmentation. Remove or bypass its meal-derived duration when source net hours are carried through. |
| `Shifts.xlsx` | Maps workers over intervals and recomputes meal-adjusted duration, including double-shift rules | Preserve row-level net values and stop recalculating roster meals. Do not smear a combined value across gaps or separate source rows. |
| `DemandIntervals.xlsx` | Repeats Demand across intervals; calculates but bypasses a meal ratio | Keep `DemandIntervalFTE = DemandShiftFTE`; remove unused meal logic after interfaces are confirmed. |
| `Demand.xlsx` | Integrates interval FTE and divides by shift duration | Retain the weighted average and reconcile it to source Demand shift FTE. |
| `AllocationByShiftAverage.xlsx` | Uses reconstructed effective/gross ratio, interval effort, standard resource FTE and actual-role FTE | Source interval FTE from roster net/clock values. Keep actual-shift and standard-resource FTE as distinctly named outputs. |
| `Allocation.xlsx` | Publishes role and resource allocation | Publish reconciled FTE without another break factor. |
| `Effort.xlsx` | Combines Demand, Allocation and Availability interfaces | Make every input's FTE basis explicit; pass actual-shift FTE to efficiency comparisons. |
| `Effort-All.xlsx` | Applies fixed `7.6 / 8 = 0.95` to capacity and a separate resource-availability branch | Do not apply the factor to already-net actual-shift FTE. Replace it only after confirming the affected input basis. |
| Capacity and Availability Redistribution | Allocate nominal availability/capacity values, often as `1.0` | Apply the approved 0.5-hour policy only when the source is nominal gross availability with no net observation. |
| `Cost..xlsx` | Converts effort with hardwired 7.6 hours | Confirm the incoming FTE definition before changing it. |

## Recommended implementation sequence

1. Approve the measure contract: `RosterNetHours`, `RosterClockHours`, `ShiftHours`, `IntervalFTE`, `ShiftFTE`, `StandardResourceFTE`, `ProductiveCareMinutes` and `DemandRosterMinutes`.
2. The user approved `Shift Net Length` for Manual Read on 16 September; Allocation Extraction already selects it. Verify source data quality and units when workbook inspection is authorised.
3. Confirm how facility-specific shift boundaries and durations are supplied. Do not join duration only by role if facility can change it.
4. Build one roster-row segmentation method that splits at configured shift and interval boundaries while preserving source row identity and total net hours.
5. Switch Allocation first. Carry the source net ratio into worker intervals and reconcile every row back to source net hours.
6. Switch the Demand historical profile to the same segmentation before role/day/shift shares are calculated. Net-field selection is already implemented in source; segmentation is not.
7. Simplify Demand conversion to actual-shift FTE from Demand roster minutes and configured shift minutes. Preserve a 7.6 equivalent only for an explicit consumer.
8. Align intervals: Demand repeats shift FTE; Allocation sums worker FTE rates. Both aggregate with the same duration-weighted-average pattern.
9. Update RosterProfile and Effort, then verify interval height and shift Demand/Allocation comparisons.
10. Update Effort-All and remove fixed 0.95 treatment from paths receiving net actual-shift FTE.
11. Handle Availability separately with the approved nominal-FTE meal policy and actual facility/role shift duration.
12. Review Cost and change/residual paths only after classifying their inputs as standard FTE, actual-shift FTE or FTE-hours.

This sequence is a proposed walk-through. It does not authorize workbook edits, synchronization or refresh.

## Required reconciliations

| Reconciliation | Required result |
| --- | --- |
| Roster row | Sum of allocated segment net hours equals source `RosterNetHours`. |
| Crossing shift | No gap or duplicate; every clock overlap belongs to exactly one configured shift. |
| Worker interval | `WorkerIntervalFTE * IntervalHours` reconstructs worker segment net hours. |
| Allocation shift | `AllocationShiftFTE * ShiftHours` equals worker net hours allocated to that shift. |
| Demand target | Allocated productive minutes reconcile to the original target. |
| Demand roster conversion | `DemandShiftFTE * ShiftMinutes` equals Demand roster minutes. |
| Demand interval round trip | Weighted interval Demand FTE equals source Demand shift FTE. |
| Facility/role duration | Every output row matches exactly one approved shift definition. |
| Downstream | Effort-All, RosterProfile and Availability apply no unlabelled extra factor. |

Test cases should include a shift below the break threshold, exactly six hours, longer than six hours, a normal eight-hour shift, overnight work, a row crossing a shift boundary, adjacent roster rows, double shifts, overlapping rows, and different facility/role durations.

## Remaining business decisions

1. Net-field selection is resolved: Manual Read now selects the user-approved `Shift Net Length`, as Allocation Extraction already does. Source-value validation and workbook parity remain pending; original `Roster Hours` is no longer used for Manual Read history.
2. Does `DemandRosterMinutes = ProductiveCareMinutes / Direct Care %` represent net working time, paid roster time including an unpaid meal span, or another capacity concept?
3. For a row crossing a shift boundary, is proportional net distribution sufficient, or must a trusted break timestamp locate the meal in a specific interval?
4. Where is the approved facility-specific shift-boundary and duration configuration maintained?
5. Which consumers genuinely require a 7.6-hour standard resource FTE rather than actual-shift attendance FTE?
6. Should Demand's interval reaggregation remain an output or become a reconciliation against the authoritative shift FTE?

## Continuation notes

- Follow repository-root `AGENTS.md` before continuing.
- Do not inspect Excel workbooks unless explicitly requested.
- Workbook-linked M-code changes require the approved source-of-truth extraction/synchronization process.
- Editing an `.m` file does not authorize synchronization, refresh, save or close operations.
- This file and the reconciliation evidence are currently uncommitted. Commit them or copy the working tree before moving to another computer.

Methods used: `data-analytics:metric-diagnostics` for numerator, denominator, grain and aggregation analysis; `data-analytics:build-report` for this durable handover.
