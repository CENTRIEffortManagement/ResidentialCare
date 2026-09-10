# Shift boundary diagnosis

## Latest update: corrected settings supplied by user

The user supplied a corrected ShiftPeriod table after the initial diagnosis. This supersedes the earlier settings reproduced below. All four roles now have AM 06:00-14:00. AIN has PM 14:00-22:00 and NIGHT 22:00-06:00. RN, AINC4 and EN have PM 14:00-22:30 and NIGHT 22:30-06:00.

Arithmetic checks using the supplied StartTime and DurationOfShifts passed for all 12 rows: each duration is positive, each end matches the next shift start modulo 24 hours, and each role totals 24 hours. The displayed EndDay fractions agree to the supplied precision. EN's previously negative durations are resolved.

With these corrected settings, the existing end-based classifier's boundary thresholds now coincide with the configured starts. This removes the settings mismatch responsible for the AIN 06:30 AM threshold; a classifier rewrite is not required solely to fix that threshold. Results still depend on the saved settings, current embedded queries, interval boundary coverage and successful downstream refresh. None of those workbook states has been verified here, and no refresh or synchronization has been performed.

The corrected durations also change Allocation FTE denominators, so values may change in addition to the boundary positions. Remaining independent concerns include midnight classification, intervals spanning reporting boundaries, consistent overnight ShiftDate assignment and Tableau's DayDate-to-Effort.Date relationship. These are not resolved by the settings arithmetic checks.

The remaining sections document the original diagnosis and earlier supplied settings for context; their boundary mismatch and EN negative-duration findings apply to the earlier table, not the corrected table.

The reviewed M source explains AM starting at 06:30 when staff start at 06:00: reporting shifts are classified using the preceding shift's **EndDay**, not the named shift's **StartDay**. The user subsequently supplied the current ShiftPeriod table, confirming AIN NIGHT EndDay = 0.270833 (06:30 to displayed precision). This review has not inspected Excel workbooks or verified that the adjacent M files match their current embedded queries or saved results.

## Supplied ShiftPeriod settings

The user-supplied settings produce these thresholds in the existing end-based classifier. Times are rounded to the nearest minute to reflect the supplied decimal precision. An actual interval must also pass the classifier's end-time conditions; missing boundary points can cause further misclassification.

| Role | Configured starts AM / PM / NIGHT | Existing classifier thresholds AM / PM / NIGHT |
| --- | --- | --- |
| RN | 06:00 / 14:00 / 22:30 | 08:15 / 14:00 / 21:15 |
| AIN | 06:00 / 14:00 / 22:00 | 06:30 / 14:15 / 21:15 |
| AINC4 | 06:00 / 14:00 / 22:30 | 07:00 / 14:15 / 21:15 |

This confirms the AIN 06:30 mechanism from the supplied settings and reviewed source; the discrepancy is not a uniform thirty-minute offset across roles or boundaries.

EN has invalid duration settings: AM = -6 hours, NIGHT = -22.5 hours, with NIGHT EndDay = -1. If EN is included in calculation, these require explicit correction or exclusion under the intended business rule; no replacement values can be inferred from this table.

The configured durations and start-to-next-start reporting windows are also distinct. For example, AIN AM has DurationOfShifts = 8.25 hours, while 06:00 to the next configured PM start at 14:00 is 8 hours. AIN NIGHT lasts 8.5 hours from 22:00 to 06:30, overlapping the next AM start by 30 minutes. RN NIGHT lasts 9.75 hours from 22:30 to 08:15. These settings should not be altered solely to force contiguous reporting windows: confirm whether durations describe attendance, paid/effective time, or FTE normalization. Actual worker attendance and reporting shift attribution require separate definitions.

## Evidence from the source

All calculation paths below are under `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/`.

- `Intervals.xlsx_PowerQuery.m:175-181`: `ShiftPeriod END` reads `Settings Data.xlsx` / `ShiftPeriod` and pivots **EndDay**, producing AM, PM and NIGHT columns by Role.
- `Intervals.xlsx_PowerQuery.m:317-330`: `Intervals` joins those role-specific end times. AM requires interval start >= NIGHT EndDay and interval end < AM EndDay. PM similarly begins at AM EndDay. Most other positive start times fall through to NIGHT.
- `Intervals.xlsx_PowerQuery.m:218-244`: staff start/end datetimes retain the input times, but the separate roster Shift label uses StartDay values taken by row position. This is a different rule from interval classification.
- `Intervals.xlsx_PowerQuery.m:301-305`: EndInterval is EndPrecise minus one minute for endpoint presentation. Duration uses EndPrecise minus StartInterval. This explains 06:29/06:30 endpoints, not a thirty-minute addition to a worker's shift.
- `Shifts.xlsx_PowerQuery.m:186-187`: worker intervals inherit the interval ShiftPeriod.
- `AllocationByShiftAverage.xlsx_PowerQuery.m:169-175,187-208`: effective effort is interval duration times the worker's effective/real duration ratio, grouped by ShiftDate, ShiftPeriod and Role; FTE is effective days times 24 divided by DurationOfShifts.
- `DemandIntervals.xlsx_PowerQuery.m:154-171`: demand intervals also inherit ShiftPeriod from Intervals. EffectiveIntervalAttendance currently equals DemandFTE; the comment explicitly identifies this as ANACC demand.
- `Demand.xlsx_PowerQuery.m:169-184,256-259`: the numerator is grouped by interval ShiftPeriod, while the duration denominator is grouped by source Shift. When those assignments disagree, the numerator and denominator can describe different time windows.
- `Effort.xlsx_PowerQuery.m:41-46,92-110`: Allocation is imported with ShiftDate renamed Date; Effort combines values at date/role/shift grain. It has no minute timestamp. Displaying these measures at individual minutes repeats shift averages; it does not calculate minute-specific staffing.

## Tableau relationships

`CLIENT/DATExx-Whiddon/UNITS/Unit1/3. Report/ShiftProfile.twb:1689-1715` confirms:

1. ResourceIntervalAllocation -> Effort: Role = Role; **DayDate = Date**; ShiftPeriod = Shift.
2. IntervalsList -> ResourceIntervalAllocation: **TimeDate = TimeDate only**.

The first relationship uses a different date convention from Allocation's ShiftDate. After midnight, DayDate and ShiftDate can differ, so this relationship can retrieve a different night's aggregate. The second does not constrain role, despite the role-expanded interval data. Validate a compound interval/role relationship against the actual row grain before changing it. Neither equality expression itself adds thirty minutes.

Tableau relationships preserve native table detail and are not equivalent to physical joins. Do not infer automatic total multiplication from a many-to-many relationship alone. See [Tableau's relationship analysis documentation](https://help.tableau.com/current/pro/desktop/en-us/datasource_multitable_analysis_overview.htm).

## Reproduction and limitations

A small PowerShell reconstruction of the exact M classification branches passed these cases using **illustrative**, unverified end settings NIGHT=06:30, AM=14:30 and PM=22:30:

| Interval endpoints | Existing classification |
| --- | --- |
| 06:00-06:29 | NIGHT |
| 06:30-06:44 | AM |
| 14:00-14:29 | AM |
| 14:30-14:44 | PM |
| 00:00-00:29 | ERROR |

This is a branch reconstruction, not Power Query execution. The midnight case shows a further conditional defect: midnight fails the positive-time NIGHT fallback. Both Intervals and Shifts also exclude time zero from their previous-day ShiftDate rule and take the morning threshold from the first ShiftPeriod row instead of the relevant role.

The prior `outputs/allocation-diagnosis/evidence.json` snapshot has 07:00 morning settings, not the screenshot's 06:30 boundary. It is historical evidence and cannot confirm the current settings. The prior refreshed-effort verification records unique date/role/shift keys, but does not validate current boundary correctness or current workbook contents. The saved Tableau worksheet state also differs from the supplied screenshot; its relationships agree with the relationship screenshot.

## Correction and validation sequence

1. Verify the exact workbook queries, Settings Data / ShiftPeriod rows and saved interval outputs under the repository's approved source-of-truth process. Inspect AIN at 05:59, 06:00, 06:29, 06:30 and its afternoon/night boundaries.
2. Establish role-specific reporting boundaries separately from actual worker attendance. If reporting AM starts at 06:00, classify from that boundary. Preserve legitimate roster overlaps and actual end times.
3. Ensure the interval grid contains every reporting boundary and midnight. Use start-inclusive/end-exclusive analytical windows and assign ShiftDate consistently by role and interval start. Do not globally subtract thirty minutes from timestamps.
4. Align Demand numerator and denominator windows; reconcile Allocation effective hours before and after reclassification. A boundary correction can move hours between shifts without changing total hours. Shift FTE changes depend on each denominator; it need not increase or decrease every shift.
5. After consistent M outputs are confirmed, relate Effort using the corresponding ShiftDate/role/shift key. Verify the role dimension and endpoint identity on the interval relationship. Consider an independent time/role scaffold for Demand so demand remains visible when no staff are allocated.
6. Reconcile overnight dates, midnight, role-specific boundaries, unique output keys and total effective hours. Compare shift averages separately from actual interval staffing.

No M source, Tableau file, workbook, refresh configuration or connection has been changed by this diagnosis. Workbook synchronization and refresh remain separate opt-in operations.
