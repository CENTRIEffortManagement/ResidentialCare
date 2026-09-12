# WorkerShiftSegments and residual effective hours

Updated 12 September 2026. Each row is one eligible worker/facility/role/date/shift. Zero-hour shifts remain available for diagnosis.

The stage measures BaselineHours, LeaveHoursInShift, UnavailHoursInShift and CombinedAbsenceHoursInShift. The combined value counts overlaps once; it need not equal the sum of the two category totals. AvailableHours is remaining **clock time** after exclusions and the daily leave decision.

ResDayShift_Calculated calculates effective hours:

- FullDayLeaveBlocked or no remaining clock time: zero.
- Entire configured shift available: ShiftDuration.
- Otherwise: max(0, min(remaining clock hours, ShiftDuration - combined distinct absence hours in the shift)).

This implements allowance-minus-absence while respecting explicit AVAIL windows. It differs from the superseded rule that merely capped remaining clock time at ShiftDuration. Absences are deducted across the configured shift; explicit AVAIL supplies an additional upper bound through remaining clock time. No break deduction or minimum-duration rule is added.

At ShiftDuration=7.6, leave 10:00–15:00 gives AM (06:00–14:00) 3.6 effective hours and PM (14:00–22:00) 6.6. The day's five leave hours do not block NIGHT. Leave 08:00–16:00 reaches the day threshold and blocks AM, PM and the entire NIGHT starting on that date.

UNAVAIL 13:00–15:00 leaves 6.6 effective hours in each adjacent shift. Two identical four-hour leave records count as four hours, not eight. Overlapping leave and UNAVAIL are also deducted once.

The public ResDayShift output still contains Role, Week, Day, Shift, EffectiveShiftHrs, Date, Name, ID and Facility-Abbrev. The downstream A.1 unit-per-row calculation has not changed.

Use [AIN_22July_HoursTest.md](AIN_22July_HoursTest.md) to compare every configured shift with independent expected hours. AvailabilityRules_TEST contains synthetic checks that also gate publication.
