# Implementation plan: Whiddon worker availability for ResDayShift

Status: plan for review. Worker Reconciliation is already in the repository's input folder. No M-code changes, synchronization, or refresh have been performed for this change.

## 1. Background and objective

ResidentialCare uses `ResDayShift` in `Capacity-ShiftAvailability.xlsx` to describe which workers can contribute capacity to each roster shift. The role-specific `CapacityDistrib(A.1)-shifts.xlsx` workbooks import its saved `Table_ResDayShift` table, match staff names to resource numbers, and match dates and shifts to periods.

The current source is the `AvailabilityMatrix` table in `StaffList Availability.xlsx`. The existing M code unpivots populated matrix cells, derives week/day/shift from column headings, removes the cell's `Available` value without testing it, and assigns the same calculated hours to every resulting row. This represents shift selections, but does not calculate availability from actual start/end timestamps or subtract leave periods.

Whiddon's `whiddon_availability_leave_extraction.xlsx` supplies dated availability, unavailability, and leave records. The requested change is to use these intervals to calculate `ResDayShift`, with a separate reconciled worker list determining who belongs in the calculation. This distinction matters because eligible workers with no availability records are to be treated as fully available.

The first implementation is temporarily restricted to **Beaudesert (BD)**. Its result will be one row for each eligible worker/role/date/shift with positive available time, including a calculated `EffectiveShiftHrs` value. Preserve the existing output interface so the rest of the workbook chain can continue to consume it.

### Current and proposed data flow

Current:

```text
StaffList Availability.xlsx / AvailabilityMatrix
  -> IMPORT AvailabilityDayShiftMATRIXRaw
  -> AvailabilityDayShift Prepare
  -> ResDayShift / Table_ResDayShift
  -> CapacityDistrib(A.1)-shifts, separately for RN, AIN and AINC4
```

Proposed:

```text
Worker Reconciliation / Employees_TABLE -> eligible BD workers
whiddon_availability_leave_extraction / Combined Output -> BD intervals
Settings Data / ShiftPeriod + PermutationDimensions -> dated shift windows
Settings Data / ShiftDuration -> standard effective-hours allowance
  -> worker rules, interval union/subtraction and shift intersections
  -> ResDayShift / Table_ResDayShift
  -> existing CapacityDistrib imports, unchanged
```

There is also a staff-identity dependency: `Availability-StaffList` feeds `StaffListMaster`, whose resource numbers are used by CapacityDistrib. The replacement must keep that staff-list interface consistent with the workers used in `ResDayShift`.

`AvailabilitiesExtracted.xlsx` was reviewed earlier and imports the same raw extraction, but it is not an additional dependency in this plan. Capacity-ShiftAvailability will read the raw extraction directly.

## 2. Scope, source of truth, and delivery boundary

Make the sole M-code edit in:

`CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m`

Use this existing `.m` file directly as the approved source. **No stale checks, workbook extraction, or source-to-workbook comparison.** Do not move the source into `Workflows`, change project configuration, or edit other client folders.

Read the existing `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Worker Reconciliation.xlsx` directly. The workbook has already been moved into the repository and renamed; no copy or move operation is required. Do not modify its contents.

Workbook synchronization, refresh, save, and close are outside this plan.

This is a source change within Capacity-ShiftAvailability, not a redesign of StaffListMaster or the later capacity-reduction stages. Do not change the RN, AIN or AINC4 CapacityDistrib M files. Do not change the raw extraction, Settings Data, or Worker Reconciliation.

**Known downstream limitation:** CapacityDistrib currently drops `EffectiveShiftHrs` and assigns `Availability = 1` to every imported row. Therefore this phase improves which shifts appear and records their available hours, but a partially available shift will still count as one availability unit downstream. The user explicitly deferred that downstream change.

## 3. Inputs and the worker population

All local workbook locations below are relative to `CLIENT/DATExx-Whiddon/UNITS/Unit1/`. Resolve imports using the target M file's existing path queries; do not add user-specific absolute paths to the M code.

| Input | Table, range, or sheet | Role in the calculation |
|---|---|---|
| `1. Input/whiddon_availability_leave_extraction.xlsx` | `Combined Output` sheet | Worker IDs, site names, date/time intervals, and availability/leave reasons |
| `1. Input/Worker Reconciliation.xlsx` | `Employees_TABLE` table | Eligible worker IDs, facilities, roles, employment types, and names where supplied |
| `2. Calculations/Settings Data.xlsx` | `ShiftPeriod` table | Shift start/end times for each role |
| Same Settings workbook | `PermutationDimensions` table | Roster dates, roles, shifts and periods |
| Same Settings workbook | `ShiftDuration` range/table, `ShiftDuration` column | Imported full-shift effective-hours allowance |
| `1. Input/1-AllocationExtracted.xlsx` | Existing `ResRolesProportion` table | Preserve the existing multiple-role naming and ancillary capping dependencies |

### Extraction preparation

Replace the `AvailabilityMatrix` import with the `Combined Output` sheet from `1. Input/whiddon_availability_leave_extraction.xlsx`.

Immediately after promoting headers:

1. Temporarily filter to `[Site Name] = "BD: Beaudesert"`.
2. Add `"Facility-Abbrev", each Text.Start([Site Name], 2), type text`.

The raw source contains `Site Name`, `Employee Name`, `Payroll Code`, `Department`, `Date From`, `Date To`, and `Availability or Leave Reason`. Treat payroll codes as text and timestamps as datetimes. The site filter is an explicitly temporary business restriction and must be commented as such.

### What the reconciliation list represents

The selected authority for worker membership is the **`Employees_TABLE` table on the `Employees-TABLE` worksheet** in `1. Input/Worker Reconciliation.xlsx`. This supplies the population to evaluate, not a declaration that every listed worker is available at every time. The extraction supplies the time constraints. Do not substitute the reconciliation workbook's `Availabilities_Employees_AVAIL` table, because it includes only explicit AVAIL records and would omit other eligible workers.

| Reconciliation column | Use in this change |
|---|---|
| `EmployeeID` | Worker identifier, matched as text to extraction `Payroll Code` |
| `Facility-Abbrev` | Facility identity and the temporary `BD` population filter |
| `Role` | Map the three approved nursing roles to RN, AIN and AINC4 |
| `Employee Roster Name` | Preferred display name; recover an unambiguous extraction name when missing |
| `EmploymentType` | Retain in worker staging/diagnostics as context; it does not determine AVAIL versus exclusion-based mode |

`Employees_TABLE` combines workers found in the availability extraction and roster, then enriches them with role/employment information from `Current_Workers` in `LBStats.xlsx`. It is not a complete independent census of current staff: workers appearing only in `Current_Workers` are not automatically added by that logic.

The earlier read-only inspection found 311 rows in this table, of which 153 were labelled BD. All employee names were blank, 74 rows lacked role/facility, and 18 rows duplicated an employee/facility/role key. These figures describe the inspected source, not fixed acceptance targets. This plan uses its saved output without refreshing its upstream queries or repairing that workbook.

### Eligibility and identity rules

Read the worker population from the local reconciliation workbook's `Employees_TABLE`:

- Retain `Facility-Abbrev = "BD"` so other facilities cannot become fully available through the no-record default.
- Match `EmployeeID` to `Payroll Code` as text, together with facility.
- Map `Registered Nurse → RN`, `Assistant in Nursing → AIN`, and `Enrolled Nurse → AINC4`.
- Exclude other roles. Report and exclude missing roles/facilities or unresolved worker identities.
- Deduplicate employee/facility/role entries. Use the reconciliation name where present; otherwise recover an unambiguous name from the extraction.
- Preserve the existing role-suffix convention for multiple-role workers.
- Report extraction workers absent from reconciliation without adding them automatically.

Treat reconciliation as the selected worker population; do not expand it with employees found only in other workforce sources.

Apply the BD filter to both worker selection and interval selection. Filtering only the intervals would incorrectly give workers from other facilities full availability under the no-record default. Workers with a missing facility must remain visible in the exclusion diagnostic rather than being assigned to BD by assumption.

## 4. Agreed availability and shift rules

### Record classification and worker mode

Classify each source record using the supplied expression unchanged:

```powerquery
if Text.Contains([Availability or Leave Reason], "UNAVAIL") then false
else if Text.Contains([Availability or Leave Reason], "AVAIL") then true
else false
```

Determine each worker/facility's rule from **all dates in the Beaudesert extraction**, before restricting calculations to the Settings roster horizon:

| Records present | Available time |
|---|---|
| Any `true` records, including mixed records | Union of explicitly available intervals minus the union of false intervals |
| Only `false` records | Entire roster horizon minus the union of false intervals |
| No records | Entire roster horizon |

Validate null/error reasons and invalid timestamps explicitly. Preserve the expression's case sensitivity.

For mixed records, explicit AVAIL windows define the limits of availability and exclusions take priority wherever they overlap. Adding a leave or UNAVAIL record must never make a worker available outside their AVAIL windows. Ordinary leave reasons classify as `false` through the final `else` branch.

Determine the mode before selecting the roster dates. For example, a worker whose `true` records all fall outside the configured roster remains restricted to those explicit windows and has no availability within that roster, even if they also have false records. Do not reclassify that worker as having no records. A false interval outside the roster cannot broaden availability inside it.

### Settings-driven shifts and hours

Use `Settings Data.xlsx` for two separate purposes:

- **Shift boundaries:** role-specific `ShiftPeriod` start/end values, checked against their durations.
- **Full-shift allowance:** replace the local calculation with `IMPORT EffectiveShiftHrs`, reading the `ShiftDuration` column's value from the `ShiftDuration` range. Validate one positive numeric value; do not hard-code 7.6.

Use `PermutationDimensions` for the roster dates, roles, and shifts. Keep `StdRosterDays` for its existing capping uses.

The inspected `ShiftPeriod` settings illustrate why the role is part of the shift key: AIN night starts at 22:00, whereas RN/AINC4 night starts at 22:30. These are examples, not constants to embed in the new code. Use `StartDay` and `EndDay` as the day-fraction endpoints, cross-check `StartTime`/`EndTimeCheck` and `DurationOfShifts`, and roll an overnight end onto the next date. Do not use the separate generic `ShiftStart` table.

The old scalar calculation `IMPORT OrdinaryWorkWk / StdRosterDays * 2` is superseded by `IMPORT EffectiveShiftHrs`. Its source is the `ShiftDuration` value, which was 7.6 in the inspected Settings workbook. Keep actual shift duration and the effective-hours allowance distinct: a fully available shift receives the imported allowance even where its clock duration differs.

### Interval calculation and aggregation

Calculate available intervals as follows:

- Union duplicate and overlapping intervals before subtraction.
- Treat starts as inclusive and ends as exclusive.
- Split availability at configured shift boundaries.
- Assign overnight shifts to their starting date and include the complete final overnight shift.
- Sum separate available periods within the same worker/facility/role/date/shift.
- Emit no row when available time is zero.

Set `EffectiveShiftHrs` to the imported allowance for a fully available shift; otherwise use the smaller of actual available hours and that allowance. Introduce no break deductions or minimum-duration rules.

Build only the shifts represented in Settings. Include source intervals that start before or end after a shift if they overlap it. Keep the complete final overnight shift, and do not repeat the extraction as a weekly or fortnightly pattern.

Full availability means the union of available segments covers the complete shift. Deduplicate and union segments before summing so overlapping records never add hours twice. Retain full precision in calculations; compare derived hour totals with a small numeric tolerance rather than rounding published hours.

### Worked acceptance examples

These examples assume AIN AM 06:00–14:00, PM 14:00–22:00, NIGHT 22:00–06:00, and an imported allowance of 7.6 hours. Tests must also verify that different Settings values change the results.

| Case | Expected result |
|---|---|
| AVAIL-only worker available 13:00–16:00 | AM row with 1 hour and PM row with 2 hours |
| Exclusion-based worker unavailable 10:00–12:00 | AM segments 06:00–10:00 and 12:00–14:00, aggregated into one 6-hour row |
| Worker has AVAIL 09:00–11:00 and UNAVAIL 10:00–12:00 | Only 09:00–10:00 remains available; one AM row with 1 hour |
| Worker has AVAIL 09:00–17:00 and UNAVAIL 12:00–13:00 | AM totals 4 hours and PM totals 3 hours; no availability outside 09:00–17:00 |
| Worker has AVAIL windows within the roster and a leave record outside it | The in-roster AVAIL windows remain the limits; the outside leave does not add availability |
| Eligible worker with no records | Every configured shift receives the imported full-shift allowance |
| AVAIL-only worker repeats the same 13:00–16:00 record | Still 1 AM hour and 2 PM hours |
| AVAIL-only worker available 23:00–02:00 next day | One NIGHT row, dated to that night's start date, with 3 hours |
| A partially available shift contains 7.8 available hours | EffectiveShiftHrs is capped at 7.6 |
| Exclusions cover the whole shift | No ResDayShift row for that shift |
| Available interval ends exactly at 14:00 | No positive-duration PM overlap from that interval |

## 5. Implementation sequence and output interfaces

1. **Use the scoped inputs.** Read `Worker Reconciliation.xlsx` from the existing `1. Input` folder through the input path query. Edit the approved adjacent M source directly, without extraction or freshness comparison. Reuse the existing input/calculation path queries.
2. **Replace matrix preparation.** Introduce the direct extraction and reconciliation imports, the BD filters, role/name preparation, and the Settings imports. Retire the matrix-unpivot path from the ResDayShift dependency chain. Replace the local effective-hours calculation and update its consumers to use `IMPORT EffectiveShiftHrs`.
3. **Calculate intervals and shifts.** Classify source records, determine worker modes, union intervals, subtract exclusions, and intersect available intervals with the role-specific shift windows. Work per worker and relevant date range rather than expanding timestamps minute by minute.
4. **Build the published rows.** Aggregate available segments, calculate effective hours, apply existing multiple-role display naming, derive Week/Day, and project the existing seven-column interface.
5. **Maintain staff and summary interfaces.** Derive Availability-StaffList from the same eligible-worker staging, independent of whether a worker has positive available time. Retain the existing ancillary cap formulas and output names; do not introduce fractional capacity weighting into them.
6. **Add diagnostics and validate.** Expose excluded/unmatched workers and required calculation checks. Gate the public ResDayShift output on successful checks, then run the focused validation described below.

Create clearly named staging queries for imports, validated workers, classified intervals, worker rules, shift windows, available segments, and aggregated results. Add the required query headers and business-rule comments.

Keep interval preparation and shift calculation separate so the transformation can be inspected query by query. Keep diagnostic worker IDs and source boundaries out of the seven-column public interface. Plan new staging queries as connection-only if workbook synchronization is separately requested later.

Preserve:

- `ResDayShift` and its intended `Table_ResDayShift` binding.
- Its seven columns: `Role`, `Name`, `Week`, `Day`, `Shift`, `Date`, `EffectiveShiftHrs`.
- Existing summary interfaces and capping rules.

Derive week numbers as `1 + Number.IntegerDivide(days since the first configured roster date, 7)` and use the existing English weekday abbreviations (`Mon` through `Sun`) for Day. Keep IDs, facilities, interval boundaries, and raw hours in staging/check queries. A NIGHT row's Date, Week and Day all describe the shift start date.

Preserve the current physical output column order and data types as well as the column names. Internally, use employee ID/facility/role/date/shift as the unique key. Before dropping IDs and facility, require that separate worker identities cannot collapse onto the same public Name/Role/Date/Shift key. Do not silently merge two different workers with the same display name.

Derive `Availability-StaffList` from eligible worker staging so workers with zero available shifts retain their staff identity.

**Do not edit CapacityDistrib.** It will continue assigning `Availability = 1` per imported row; partial hours will be available in ResDayShift but will not yet affect downstream capacity weighting.

## 6. Validation and acceptance

Add `ResDayShift_CHECK` and worker-exclusion diagnostics within the same M file. Block output for invalid source intervals/settings, ambiguous published identities, duplicate output keys, or failed hour reconciliation.

Separate expected exclusions from calculation failures. A non-nursing title, worker absent from the selected reconciliation list, or unresolved eligibility field belongs in the diagnostic output. Invalid required Settings, invalid intervals for included workers, or ambiguous identities in published output must prevent publication. Checks must read the calculated staging result, not the gated public ResDayShift query, to avoid a dependency cycle.

Validate:

- AVAIL-only, exclusion-only, mixed, and no-record workers.
- Beaudesert filtering on both inputs and all approved role mappings.
- Exact reason classification, duplicate intervals, and separated available periods.
- Boundary crossings, overnight shifts, role-specific night starts, and roster endpoints.
- Full-shift allowance, partial-hour caps, and zero availability.
- Worker exclusions, name recovery, and output uniqueness.

Run available M syntax validation and focused read-only regression calculations, including the worked examples above. Keep verification support temporary rather than creating additional product source files. Review query headers, dependency order, final `in` targets, and preserved interfaces.

Acceptance requires:

- No dependency on AvailabilityMatrix in the ResDayShift calculation.
- Only eligible BD nursing worker identities in the calculated population; every excluded or unmatched category is inspectable.
- Valid, unique role/shift settings for every published period, with complete overnight handling.
- No double-counted available minutes, no available segment intersecting a false interval for any worker, and no output outside the configured shift windows.
- Workers with any true records have no available segments outside the union of their explicit AVAIL windows, including when false records exist elsewhere in the extraction.
- Exactly one positive-hours row per eligible worker/role/date/shift, with the imported allowance and partial-hour cap applied correctly.
- An unchanged seven-column ResDayShift interface and consistent names/roles in Availability-StaffList.
- No edits to CapacityDistrib, Settings Data, StaffListMaster, or other client folders.

Parser validation proves syntax only. Independent interval regression checks support the arithmetic but do not prove execution in Excel. Report those verification limits explicitly; live workbook refresh testing remains outside this source-only implementation.

## 7. Handoff and deferred work

Deliver only the scoped M-source change. Report validation results and any unresolved worker exclusions; do not synchronize or refresh workbooks.

The revised plan document is the planning artifact, not evidence that these implementation steps are complete. A later explicit sync request must name the exact workbook/source pair and preserve its query/table bindings. A later refresh request is a separate operation. Making CapacityDistrib consume partial EffectiveShiftHrs, repairing reconciliation completeness, and redesigning StaffListMaster identity handling are deferred work and must not be folded into this change.
