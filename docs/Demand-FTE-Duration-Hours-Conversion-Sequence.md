# Demand FTE, Duration and Hours Conversion Sequence

## Purpose

This note records the active Unit1 demand calculation from standard-FTE inputs through interval demand effort. It distinguishes configured standard shift hours, actual shift duration and interval duration so meal or effective-hours adjustments are not applied twice.

Companion diagram: [Demand-FTE-Duration-Hours-Conversion-Sequence.mmd](Demand-FTE-Duration-Hours-Conversion-Sequence.mmd).

## Duration definitions

| Measure | Authority | Meaning |
| --- | --- | --- |
| `ShiftDuration` | `Settings Data.xlsx` | Standard shift hours: configured net roster hours represented by one standard-FTE unit, excluding meal time. |
| `StartTime` and `EndTime` | `2-DemandExtract.xlsx` | Dated shift endpoints constructed from the Settings `StartDay` and `EndDay` values. |
| `DurationOfShifts` | `Settings Data.xlsx` → `ShiftPeriod` | Elapsed shift duration in hours; Demand Extract validates it against `Duration.TotalHours(EndTime - StartTime)`. No meal time is deducted. |
| `Duration` | `Intervals.xlsx` | Elapsed width of one interval, stored as a fraction of a day. |
| `ShiftDuration` in `DemandIntervals.xlsx` | Renamed imported `DurationOfShifts` | Actual shift duration, not the standard-FTE setting. |

## Complete file sequence

### `Settings Data.xlsx`

Supplies the two duration authorities used by the demand path:

```text
ShiftDuration = configured standard net roster hours per FTE unit, excluding meals
DurationOfShifts = actual shift duration in hours = EndTime − StartTime
```

No FTE conversion occurs in this workbook for this sequence.

## Period and interval granularity

| File | Calculation grain | Period meaning |
| --- | --- | --- |
| `Demand-MasterRoster Manual Read.xlsx` | Facility, role, fortnight day and shift | A complete AM, PM or NIGHT target cell; no interval rows. |
| `2-DemandExtract.xlsx` | Facility, role, planning date and shift period | One complete dated shift from `StartTime` to `EndTime`. `Period` identifies the dated shift cell. |
| `Intervals.xlsx` | Role and `IntervalID` | A smaller variable-width slice between adjacent time boundaries. Multiple intervals can make up one shift. |
| `DemandIntervals.xlsx` | Facility, role, date, shift and interval | Repeats the shift-level demand across every interval covered by that shift and calculates `DemandEffort` per interval. |
| `Demand.xlsx` | Facility, role, date and shift period for `ShiftDemandHCAverageANACC` | Aggregates interval calculations back to the shift-period comparison grain. |

In this sequence, a **shift period** means the complete configured AM, PM or NIGHT shift. An **interval** is a smaller time slice inside or across shift periods, created wherever demand or roster timestamps introduce a boundary.

### `Demand-MasterRoster Manual Read.xlsx`

Publishes target roster demand as standard-shift equivalents:

```text
RedistributedRosterFTE = target roster hours / configured ShiftDuration
```

The published FTE is therefore denominated in configured standard-shift hours.

### `2-DemandExtract.xlsx`

Restores target roster hours from the standard-FTE publication:

```text
DemandHRS = RedistributedRosterFTE × configured ShiftDuration
```

It then distributes those hours over the actual shift duration:

```text
DemandFTE = DemandHRS / DurationOfShifts
```

Demand Extract constructs `StartTime` and `EndTime`, then requires `DurationOfShifts` to equal `Duration.TotalHours(EndTime - StartTime)`. This is elapsed time with no meal deduction. Target hours and configured `ShiftDuration` already exclude meals, so the current `DemandFTE` calculation changes standard FTE into average attendance across the actual shift duration. It must not be treated as standard FTE in downstream comparisons.

### `Intervals.xlsx`

Creates the time boundaries used to divide demand shifts:

```text
Duration = EndPrecise − StartInterval
```

`Duration` is stored as a fraction of a day. This workbook does not assign demand FTE or demand effort.

### `DemandIntervals.xlsx`

Imports `DemandFTE`, `DemandHRS` and `DurationOfShifts` from Demand Extract, then renames `DurationOfShifts` to `ShiftDuration` locally.

The active interval calculation is:

```text
EffectiveIntervalAttendance = DemandFTE
DemandEffort = interval Duration × 24 × DemandFTE
```

Multiplication by 24 converts interval duration from days to hours. Because `DemandFTE` already contains the standard-hours-to-actual-duration relationship, `DemandEffort` inherits that adjustment exactly once.

`EffectiveDuration` and `IntervalEffectiveRatio` are calculated but are not used by the active AN-ACC demand calculation. The alternative `DemandFTE × IntervalEffectiveRatio` expression is commented out. Applying it in addition to the active Demand Extract conversion would introduce a second reduction.

The target-hour reconciliation is:

```text
SUM(interval Duration × 24) = actual DurationOfShifts
SUM(DemandEffort) = DemandHRS
```

### `Demand.xlsx`

The AN-ACC branch imports interval attendance and duration, then calculates:

```text
UnitIntervalEffort = EffectiveIntervalAttendance × interval Duration
ShiftDemandHCAverage = SUM(UnitIntervalEffort) / shift duration in days
```

The day units cancel in the final division, returning average shift attendance. No additional meal adjustment is applied.

## Meal/effective-hours boundary

Target hours and configured `ShiftDuration` already exclude meals. Dividing `DemandHRS` by `DurationOfShifts` where that duration includes the meal window produces average attendance across the actual shift duration, not standard FTE. `DemandIntervals.xlsx` carries that average attendance into interval effort; it must not also apply `IntervalEffectiveRatio`, which would reduce for meals again.

## Target-hour invariants

At each facility, role, date, shift, start-time and end-time grain:

```text
DemandFTE × DurationOfShifts = DemandHRS
SUM(interval hours) = DurationOfShifts
SUM(interval DemandEffort) = DemandHRS
```

Checks should reject missing or duplicate interval joins, gaps or overlaps, null or non-finite measures, and residuals outside the approved tolerance. Legacy `/2` and `× 0.936` diagnostic approximations are not part of this conversion sequence.

## File relationship summary

The primary demand flow is:

```text
Demand-MasterRoster Manual Read.xlsx
    → 2-DemandExtract.xlsx
    → DemandIntervals.xlsx
    → Demand.xlsx
```

`Intervals.xlsx` is a supporting timing service. Demand Extract contributes demand `StartTime` and `EndTime` boundaries; Intervals returns interval IDs, endpoints and widths to DemandIntervals and the interval timeline to Demand.

The principal interfaces are:

| Provider | Interface | Consumer | Contribution |
| --- | --- | --- | --- |
| `Demand-MasterRoster Manual Read.xlsx` | `MinuteWorkersFTE_TABLE`, `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE`, `MinuteWorkersFTE_CHECK` | `2-DemandExtract.xlsx` | Target standard FTE, paired profile and upstream validation. |
| `2-DemandExtract.xlsx` | `ShiftUnitDemandHRS` | `DemandIntervals.xlsx` | Dated shift demand, target hours, current attendance, shift endpoints and actual shift duration. |
| `2-DemandExtract.xlsx` | `StartTime`, `EndTime` in `ShiftUnitDemandHRS` | `Intervals.xlsx` | Demand shift boundaries used in the common timing grid. |
| `Intervals.xlsx` | `RoleIntervalID`, `Intervals` | `DemandIntervals.xlsx` | Interval identifiers, endpoints, role/shift classification and interval duration. |
| `DemandIntervals.xlsx` | `Table_ShiftDemandUnitINTERVAL` | `Demand.xlsx` | Shift demand expanded to interval grain with attendance and interval effort. |
| `Intervals.xlsx` | `IntervalsList` | `Demand.xlsx` | Timeline used to attach the calculated shift average to interval endpoints. |

## Reusable measure-lineage and conversion audit

This demand review used the following process. Apply the same process to Allocation.

1. **Set the exact file chain.** Identify the primary business flow separately from supporting settings, timing and lookup files.
2. **Name every published interface.** Record the table, worksheet or query passed across each workbook boundary.
3. **Define every measure before using it.** State its unit and business meaning, including whether hours are target, net, elapsed, effective or interval hours.
4. **Record the grain at every stage.** Examples include source row, worker/shift, role/date/shift, and role/date/shift/interval.
5. **Expand every conversion formula.** Show the numerator, denominator and unit conversion, including factors such as `× 24` or `/ 60`.
6. **Trace renames and overloaded names.** Identify when the same measure changes name or one name represents different measures in different workbooks.
7. **Separate active calculations from diagnostics.** Commented-out expressions, unused ratios and legacy checks must not be described as active business logic.
8. **Locate each meal or effective-hours adjustment.** Establish where it first occurs and prevent the same adjustment from being applied again downstream.
9. **Test boundary reconciliations.** Reconstruct the upstream measure from the downstream rows at the same business key.
10. **Check joins and grain preservation.** Reject missing matches, duplicate matches, gaps, overlaps and unintended row multiplication.
11. **Classify the final measure correctly.** Do not compare standard FTE, average attendance, hours and interval effort as though they share one basis.
12. **Draw the chart last.** Use solid arrows for the primary business flow, dotted arrows for supporting inputs, bold query/interface names, and `[bracketed commentary]` for business interpretation and cautions.

## Allocation audit starting point

Apply the audit to this provisional chain:

```text
1-AllocationExtracted.xlsx
    → Intervals.xlsx and Shifts.xlsx
    → AllocationByShiftAverage.xlsx
    → Allocation.xlsx
    → Effort.xlsx
```

Start by establishing:

- whether the source `Hours` field is net roster hours and where meal time has already been removed;
- the distinction between roster-row duration, interval duration, effective duration, actual shift duration and configured standard shift hours;
- whether interval effort sums back to each source roster row's net hours;
- whether worker interval effort aggregates uniquely to worker/date/shift and role/date/shift;
- which outputs divide by actual shift duration and which divide by configured `ShiftDuration`;
- whether `ResourceShiftAllocation`, `RoleShiftAllocation`, `RoleShiftFTE` and the published Allocation measures share the same denominator;
- whether Allocation and Effort preserve those measure definitions without another meal or duration adjustment.
