# DATExx-Whiddon capacity shift availability rules

This document explains how the authoritative Power Query `.m` source determines which workers are available for each roster shift and how many effective hours they contribute. It describes the source implementation as of 19 September 2026. The four-worker refresh matched all 336 independently calculated day/shift rows. The sample filters were then removed from the source; full BD and downstream Capacity results are under review after a user-run Unit1 refresh.

Source: [Capacity-ShiftAvailability.xlsx_PowerQuery.m](../CLIENT/DATExx-Whiddon/UNITS/Unit1/2.%20Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m).

## 1. Sources and scope

All workbook paths below are relative to `CLIENT/DATExx-Whiddon/UNITS/Unit1/`.

| Source | Table or sheet | Purpose |
| --- | --- | --- |
| `1. Input/Worker Reconciliation.xlsx` | `Employees_TABLE` table | Authoritative population of workers for this calculation, including identity, role and facility |
| `1. Input/whiddon_availability_leave_extraction.xlsx` | `Combined Output` sheet | Availability and exclusion records, including payroll code, dates, times and reason |
| `2. Calculations/Settings Data.xlsx` | `PermutationDimensions` table | Roster dates, roles, shifts and period identifiers |
| `2. Calculations/Settings Data.xlsx` | `ShiftPeriod` table | Role-specific shift start and end boundaries |
| `2. Calculations/Settings Data.xlsx` | `ShiftDuration` table or named range | Effective hours for a full shift and the threshold for a full day of recognised leave |
| `1. Input/1-AllocationExtracted.xlsx` | `ResRolesProportion` table | Existing multiple-role display-name convention |

The temporary facility restriction is exactly `Site Name = "BD: Beaudesert"`. It is applied after promoting the extraction sheet's headers. `Facility-Abbrev` is the first two characters of `Site Name`, giving `BD`. Reconciled workers must also belong to `BD`.

The old `AvailabilityMatrix` is no longer the availability source.

## 2. Which workers are included?

Membership comes from `Employees_TABLE`, not from the availability extraction alone. The calculation does not add workers who are absent from reconciliation, even if they appear in the extraction.

The matching key is reconciled `EmployeeID` plus `Facility-Abbrev`, matched to extraction `Payroll Code` plus `Facility-Abbrev`. Identifiers are normalized as text. The role mappings are:

| Reconciled role | Calculation role |
| --- | --- |
| Registered Nurse | RN |
| Assistant in Nursing | AIN |
| Enrolled Nurse | AINC4 |

Missing IDs, missing facilities, facilities outside BD, missing roles and other roles are excluded. Employment type is retained during preparation but does not independently determine eligibility.

The preferred name is `Employee Roster Name` from reconciliation. If there is no reconciled name, the calculation uses the extraction's `Employee Name` only when it resolves to one distinct name for the worker/facility. Missing or ambiguous names exclude the worker. Duplicate reconciliation entries are consolidated.

Workers with multiple reconciled roles, or covered by the existing multiple-role naming convention, receive names such as `Worker Name (RN)`. A worker can have a separate availability row for each eligible role. Published names must remain unique because existing downstream calculations still use names as keys.

`WorkerAvailability_DIAGNOSTICS` reports excluded reconciliation records and extraction workers absent from reconciliation. Inclusion in reconciliation is not, by itself, proof of available hours.

## 3. Classify each availability or leave record

Classification uses `Availability or Leave Reason`, in this order:

| Reason text | Treatment |
| --- | --- |
| Contains uppercase `UNAVAIL` | Timed unavailability; subtract only its recorded overlap |
| Otherwise contains uppercase `AVAIL` | Explicit availability |
| Matches a recognised leave description | Leave; subtract its overlap and count it toward the daily leave threshold |
| Any other reason or blank | Unknown; validation failure |

Matching is case-sensitive. `UNAVAIL` must be checked first because it also contains `AVAIL`. A null or invalid reason is a validation failure, not an unrestricted availability instruction.

Both timestamps must be valid and the end must be later than the start. Intervals include their start and exclude their end: an absence ending at 14:00 does not overlap a shift beginning at 14:00.

## 4. Decide the worker's baseline availability

The calculation starts afresh on each calendar day. An AVAIL record restricts only the dates its interval overlaps. It does not restrict availability on other dates.

| Worker records on a calendar day | Baseline and resulting availability on that day |
| --- | --- |
| At least one explicit AVAIL interval overlapping this date | Available only within the union of this date's AVAIL windows, minus this date's exclusions |
| Exclusion records only | Start with the whole day available, then subtract this date's exclusions |
| No records overlapping this date | Whole day available, provided the worker is eligible and has a resolved name |

Exclusions always take precedence on the dates they affect. Adding a leave record to a date with explicit AVAIL windows does not make the worker available outside those windows on that date.

For example, AVAIL from 06:00 to 14:00 on 22 July restricts 22 July to those hours. On 23 July the worker starts fully available unless that date has its own AVAIL or exclusion intervals. An AVAIL outside the roster calendar dates has no effect on the roster.

`AvailabilityDailyWindows` constructs these daily baselines, which `WorkerAvailabilityRules` publishes as `AvailableIntervals`. Even workers with no records receive baseline intervals covering the roster.

Midnight is an exclusive endpoint. AVAIL ending at midnight does not restrict the following day. A recorded AVAIL spanning midnight affects both dates for the portions actually recorded, and a night shift uses the rules of each calendar date it crosses.

## 5. Apply recognised leave and timed UNAVAIL

The full-day threshold is the positive, finite `ShiftDuration` value from Settings Data, read by `EXTRACT EffectiveShiftHrs`. Examples below use 7.6 hours; the calculation uses the setting rather than a hard-coded threshold.

For each worker and facility:

1. Split recognised leave intervals at calendar midnight where necessary.
2. Combine duplicate and overlapping leave intervals within each date.
3. Sum distinct leave hours for each date. Overlaps count once.
4. If that date's leave hours reach `ShiftDuration`, block every shift **starting** on that date, including NIGHT.
5. Otherwise, subtract only the recorded leave and UNAVAIL overlap from each shift's baseline. Union leave and UNAVAIL before subtraction so overlapping records are counted once.

`UNAVAIL` never contributes to the daily leave threshold. A midnight-to-midnight UNAVAIL record removes the time it covers, but it does not automatically block a NIGHT shift's hours after midnight. A midnight end is exclusive and does not add an extra day.

Complete boundary dates are considered before clipping to the roster's shift hours. This prevents leave before the first shift's start from being omitted from that date's total. Totals are calculated per worker/facility, not multiplied by the worker's number of roles.

| Recorded source on one date | Distinct leave hours | Result with ShiftDuration = 7.6 |
| --- | --- | --- |
| 00:00 to next-day 00:00 UNAVAIL | 0 | Recorded time only; next-day NIGHT portion remains possible |
| 08:00–16:00 maternity leave | 8 | All shifts starting on that date blocked |
| 06:00–14:30 unpaid leave plus identical study leave | 8.5, not 17 | All shifts starting on that date blocked |
| 06:00–10:00 plus 12:00–16:00 recognised leave | 8 | All shifts starting on that date blocked |
| Two identical 06:00–10:00 leave records | 4, not 8 | Only recorded shift overlap deducted |
| 10:00–12:00 UNAVAIL | 0 | Only recorded shift overlap deducted |

The full-day leave decision follows the shift's start date. A NIGHT shift starting on a full-day leave date is blocked in full. A NIGHT shift starting on another date uses the recorded interval overlaps, including after midnight.

## 6. Apply the configured shift boundaries

`PermutationDimensions` supplies the roster's role/date/shift combinations. `ShiftPeriod` supplies the corresponding role-specific boundaries. The calculation does not assume all roles have identical shift times.

For each eligible worker and configured shift:

1. Intersect the worker's baseline availability with that shift's start and end.
2. Subtract the union of actual leave and UNAVAIL intervals, unless full-day leave blocks the starting date.
3. Retain the remaining disjoint available segments.
4. Sum those segments once to obtain available clock hours.

Availability crossing a shift boundary contributes separately to each affected shift. For example, 13:00–16:00 availability across a 14:00 boundary contributes one clock hour to the earlier shift and two to the later shift, before any exclusions.

Overnight shifts retain the date on which they start. Their end is on the following date, including the final roster night's end. Week numbers start at 1 using the earliest configured roster date; day labels come from the shift's start date.

## 7. Calculate EffectiveShiftHrs

| Remaining availability within a shift | EffectiveShiftHrs |
| --- | --- |
| Covers the complete configured shift | `ShiftDuration` |
| Covers only part of the shift | Greater of zero and the smaller of remaining clock hours and `ShiftDuration` minus distinct shift absence hours |
| No remaining available time | No published ResDayShift row |
| Full-day leave on the shift's start date | Zero; no published ResDayShift row |

Several separate available segments within a shift produce one row and one summed hours value. Overlaps cannot increase the total.

A complete shift receives the standard allowance even when its clock duration differs from that allowance. For example, a fully available 7.5-hour shift receives 7.6 effective hours when `ShiftDuration` is 7.6. There are no additional break deductions or minimum-duration rules.

## 8. Validation and published outputs

`ReconciledWorkers_CHECK` checks worker-key and published-name uniqueness. It gates `Availability-StaffList`, which includes eligible workers even when they have no available shifts. Staff-list publication does not require calculation of every shift.

`ResDayShift_CHECK` includes the worker checks, known-answer availability-rule tests, unique daily leave totals, settings and role coverage, required source records, unique worker/shift rows, segment boundaries, exclusions, disjointness, hours, facility and roles. Invalid records for eligible workers, and records without a usable payroll code, block shift publication. Invalid records tied only to ineligible workers do not enter the eligible calculation.

Failed required checks stop `ResDayShift` from publishing results. Zero-hour rows remain available internally for checks but are omitted from the public output.

`ResDayShift` publishes one positive row per worker/facility/role/date/shift, with these columns in order:

| Column | Meaning |
| --- | --- |
| Role | RN, AIN or AINC4 |
| Week | Roster week number |
| Day | Day-of-week label |
| Shift | Configured shift label |
| EffectiveShiftHrs | Effective available hours under the rules above |
| Date | Shift start date |
| Name | Resolved display name, including role suffix where required |
| ID | Reconciled EmployeeID, stored as text |
| Facility-Abbrev | Facility identifier, stored as text |

The workbook interface expected by Capacity Distribution is `Table_ResDayShift` in `Capacity-ShiftAvailability.xlsx`. The query name is `ResDayShift`; the expected loaded Excel table name is `Table_ResDayShift`.

## 9. Downstream capacity interpretation

All three role versions of `CapacityDistrib(A.1)-shifts` read `Table_ResDayShift` and filter it by role. Their current calculation selects the columns it needs and continues to match workers by name; the appended ID and facility columns are available for future use.

A.1 currently assigns availability of **1 per published worker/shift row**. It does not yet weight capacity using `EffectiveShiftHrs`. Consequently, a short partial shift still contributes one availability unit downstream. Fractional-hour capacity weighting is a separate, deferred change.

A.2 reads outputs from A.1. B reads outputs from A.1 and A.2. Neither directly reads other tables from Capacity-ShiftAvailability.

The retained legacy `RoleResDayAvailabilityCapped` and `RoleAvailabilityCapped` branch does not determine `ResDayShift`. Its `StdRosterDays` parameter does not set effective shift hours or the full-day leave threshold.

## 10. Execution and maintenance notes

Only eligible workers' intervals are grouped for the shift calculation. Availability baselines reset per calendar date, and interval processing is limited to relevant roster calendar dates. Prepared workers, selected Settings tables, reusable interval lists, scalar shift measurements and small check results are buffered where reused. Historical interval lists are not retained in the expanded scalar shift output.

Buffers apply within a query evaluation; they do not guarantee a shared cache across separately refreshed outputs. Source rule tests and syntax validation do not establish Excel refresh performance.

Changing the M source does not synchronize it to Excel or refresh the workbook. Synchronization and refresh remain separate operations under the repository's approved process.
