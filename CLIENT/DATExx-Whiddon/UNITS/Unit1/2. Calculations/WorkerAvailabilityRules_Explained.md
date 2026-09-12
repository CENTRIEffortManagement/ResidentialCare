# WorkerAvailabilityRules: daily baselines and separate absences

Updated 12 September 2026 for the revised source. The earlier copied inspection query has been retired because it combined UNAVAIL and leave and promoted both to whole days.

1. AvailabilityEligibleRecords retains valid records for eligible worker/facility keys over complete roster calendar dates. Multiple roles do not duplicate source records.
2. WorkerLeaveDays reads **Leave records only**, splits at midnight, counts distinct covered hours per date and compares them with EXTRACT EffectiveShiftHrs (Settings ShiftDuration). Its scalar FullDayLeaveBlocked flag applies to every shift starting on the date, including the complete NIGHT shift. UNAVAIL never contributes.
3. WorkerAvailabilityRules groups separate AvailableIntervals, UnavailIntervals and LeaveIntervals per worker/facility, then attaches them to every eligible role. AvailableIntervals includes full days when there is no AVAIL on that date. If AVAIL exists, only that date is restricted. Workers without records receive full baseline availability.
4. WorkerShiftSegments joins each role's configured shifts and each date's leave decision. It subtracts only the recorded overlap of UNAVAIL and partial leave. Overlapping categories are unioned so time is not deducted twice.
5. ResDayShift_Calculated applies the effective allowance after these measurements. Zero-hour shifts remain visible here; ResDayShift publishes positive rows only.

**UNAVAIL never expands to an entire shift merely because it overlaps that shift.** Nor does a zero-hour AM shift automatically remove PM or NIGHT. Recognised daily leave reaching ShiftDuration is what blocks all shifts on the date.

HasAvailability and the combined ExcludedIntervals field are retired. The separate interval categories and WorkerLeaveDays decision now carry the business meaning directly. Historical queries depending on those retired fields must be replaced with [the revised test](AIN_22July_HoursTest.md).
