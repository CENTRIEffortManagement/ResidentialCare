# How the four-worker ResDayShift result is calculated

This walkthrough follows the authoritative Availability Power Query and the independent 336-row `day_shift_matrix_expected.csv`. `day_shift_matrix_reconciled.csv` adds the **actual refreshed** `WorkerShiftSegments` measurements and `ResDayShift` result for every row. Its `SavedWorkerShiftSegmentsHours` and `PresentInSavedResDayShift` columns refer to the earlier, pre-repair workbook save; use the `Actual...` columns for the current result.

The saved `Capacity-ShiftAvailability.xlsx` refresh at 16:32 on 19 September 2026 contains 336 segment rows and 108 published rows. All 336 rows match the independent matrix, and all 12 saved `ResDayShift_CHECK` rows pass. Read-only re-extraction confirms its embedded Power Query matches the adjacent authoritative `.m`. The comparison used workbook SHA-256 `8989706BB54804A412C7969E9165CD81DDA47ED95A29214F00ADB47429566846`.

## One row at a time

1. `ReconciledWorkers_Eligible` selects the BD worker identity and nursing role. `AvailabilityShiftWindows` supplies the date and the role's AM, PM or NIGHT start and end time. The sample has 28 dates × 3 shifts × 4 workers = 336 candidate rows.
2. `AvailabilityRecords` classifies each source interval as `AVAIL`, `UNAVAIL`, recognised leave or invalid. The four sampled workers have no `AVAIL` records, so their starting window is the full calendar day. Any future `AVAIL` record would restrict only the dates and times it covers.
3. `WorkerLeaveDays` unions overlapping recognised-leave intervals on each calendar date. At least 7.6 distinct leave hours blocks all shifts **starting** on that date. `UNAVAIL` never contributes to this daily threshold.
4. `WorkerShiftSegments` clips the worker's `UNAVAIL` and leave intervals to each actual shift. It unions overlapping absences, subtracts them from the baseline, and records remaining clock hours. A NIGHT shift crosses midnight, so intervals on both dates can affect it.
5. `ResDayShift_Calculated` assigns 7.6 effective hours to a fully available shift. For a partly available shift, it uses the smaller of remaining clock hours and `7.6 − distinct absence hours`, with a floor of zero. Full-day recognised leave gives zero.
6. `ResDayShift_CHECK` validates the rows. `ResDayShift` publishes only rows with `EffectiveShiftHrs > 0`. Presence means positive availability for that worker, date and shift; the final Capacity calculation occurs downstream.

## Example: Monday 20 July 2026

All four workers start with an eight-clock-hour shift baseline. AM is 06:00–14:00, PM is 14:00–22:00 and NIGHT is 22:00–06:00 the next day. The configured effective shift allowance is 7.6 hours.

| Worker | Source intervals and day rule | AM | PM | NIGHT |
| --- | --- | ---: | ---: | ---: |
| 19822 | `UNAVAIL` 06:00–22:00 (#3600). No recognised leave that day. | 0 | 0 | **7.6** |
| 19835 | `UNAVAIL` all day (#3708); PEL leave 14:00–22:00 (#3709) gives 8 distinct leave hours, so the day's shifts are blocked. | 0 | 0 | 0 |
| 24562 | `UNAVAIL` 00:00–14:00 and 22:00–24:00 (#1943–1944). The gap is the PM shift. | 0 | **7.6** | 0 |
| 25376 | `UNAVAIL` 00:00–06:00 and 22:00–24:00 (#1534–1535). AM and PM have no overlap. | **7.6** | **7.6** | 0 |

Bold values are the rows that appear in the refreshed `ResDayShift` for this date. For example, worker 24562's 16 hours of `UNAVAIL` do **not** wipe out the uncovered PM shift: the daily full-shift threshold applies to recognised leave only. Worker 19822's NIGHT result is 7.6 because the 06:00–22:00 `UNAVAIL` ends exactly when NIGHT starts.

## Refreshed sample totals

| Worker | Candidate shifts | Published rows | Published effective hours |
| --- | ---: | ---: | ---: |
| 19822 | 84 | 25 | 190.0 |
| 19835 | 84 | 5 | 38.0 |
| 24562 | 84 | 32 | 243.2 |
| 25376 | 84 | 46 | 349.6 |
| **Total** | **336** | **108** | **820.8** |

Filter `day_shift_matrix_reconciled.csv` by `EmployeeID`, `Date` and `Shift` to see each source record, the recognised-leave decision, hours of `UNAVAIL` and leave inside the shift, residual hours, effective hours, and the actual publication result. `MatchesRefreshedResult` is `True` on all 336 rows. Run `compare_refreshed_availability.ps1` after another save to repeat the read-only comparison.
