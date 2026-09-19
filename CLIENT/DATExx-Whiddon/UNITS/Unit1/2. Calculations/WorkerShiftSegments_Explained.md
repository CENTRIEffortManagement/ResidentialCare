# WorkerShiftSegments explained

The authoritative calculation is in `Capacity-ShiftAvailability.xlsx_PowerQuery.m`. The four-worker refresh matched the independent day/shift matrix on all 336 rows. The sample filters have since been removed from the source; the full BD result is under review after a user-run Unit1 refresh.

`WorkerShiftSegments` makes one row for each eligible worker and configured shift, including shifts that end with zero hours. It joins the worker's interval lists from `WorkerAvailabilityRules` to the leave decision in `WorkerLeaveDays` for the shift's **start date**.

For each shift it:

1. Finds the AVAIL baseline inside the shift. A date without AVAIL starts fully available.
2. Finds the recorded leave and UNAVAIL times that overlap the shift. An interval ending exactly at shift start does not overlap it.
3. Combines leave and UNAVAIL before subtracting them, so overlapping records are counted once.
4. Sets the remaining time to zero if recognised leave reached the whole-day threshold on the shift's start date. This also blocks a NIGHT shift starting on that date.
5. Measures `AvailableHours`, checks whether the entire shift remains (`FullShift`), and validates the remaining segments (`SegmentsValid` and `Disjoint`).

`ResDayShift_Calculated` then calculates `EffectiveShiftHrs`. A complete shift gets the Settings `ShiftDuration` allowance. A partial shift gets the greater of zero and the smaller of remaining clock hours and the allowance after distinct shift absence hours. A zero result does not appear in the published `ResDayShift` table.

For a missing shift, filter `WorkerShiftSegments` by worker, `Date` and `Shift`. Read `FullDayLeaveBlocked` first, then `BaselineHours`, `LeaveHoursInShift`, `UnavailHoursInShift`, `CombinedAbsenceHoursInShift` and `AvailableHours`. The matching `ResDayShift_Calculated` row gives the final decision.
