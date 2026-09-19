# WorkerAvailabilityRules explained

The authoritative calculation is in `Capacity-ShiftAvailability.xlsx_PowerQuery.m`. The four-worker refresh matched the independent day/shift matrix on all 336 rows. The sample filters have since been removed from the source; the full BD result is under review after a user-run Unit1 refresh.

`WorkerAvailabilityRules` prepares three lists for each eligible worker:

| List | Meaning |
| --- | --- |
| `AvailableIntervals` | Hours offered by AVAIL records on each date. If a date has no AVAIL record, the whole date is the starting baseline. |
| `UnavailIntervals` | Recorded UNAVAIL times. These are not expanded to a whole day. |
| `LeaveIntervals` | Recorded times for recognised leave. |

The query takes valid source rows from `AvailabilityEligibleRecords`, groups them by worker and facility, then attaches the lists to each eligible worker and role. An AVAIL record on one date does not restrict another date. A worker with no matching records gets a full roster baseline and empty UNAVAIL and Leave lists.

This query does not subtract hours or decide whole-day leave. `WorkerLeaveDays` separately adds up distinct recognised leave hours for each date. `WorkerShiftSegments` uses both results to measure each configured shift.

To inspect a worker, filter `WorkerAvailabilityRules` by `EmployeeID` and `Facility-Abbrev`. Check the three lists, then check `WorkerLeaveDays` for the shift's start date. `WorkerShiftSegments` and `ResDayShift_Calculated` show the shift result.
