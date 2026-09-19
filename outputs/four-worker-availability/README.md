# Four-worker availability walk-through

For the calculation steps and a 20 July example, see `availability-walkthrough.md`. The complete worker/date/shift evidence and the saved post-refresh comparison are in `day_shift_matrix_reconciled.csv`: all 336 expected rows match actual measurements and publication, and all 12 saved `ResDayShift_CHECK` rows pass.

After that bounded review, the three four-worker sample filters were removed from the authoritative Availability `.m` source. At the last verified source-edit handoff, the saved workbook still contained the four-worker version; the full BD population had not yet been refreshed or validated.

The user subsequently reported running Unit1 and is checking its results. The generated note for run `20260919-165249-04a88aab` reports a completed technical refresh, with business reconciliation unevaluated. The statement above describes the last verified source-edit handoff; this directory has not yet reconciled the later run's full BD or downstream Capacity outputs.

The `.m` files are authoritative. The existing Availability and AIN B `.m` sources contain the revised rule and the A.2-to-A.1 lineage gate. An intermediate Excel refresh showed a `FilePathUrl` error before shift calculation; the `.m` resolvers now accept the workbooks' existing named cells. The Availability workbook was subsequently refreshed and saved, and its four-worker result matches the independent expected matrix. AIN B and downstream Capacity still need their ordered rerun. `day_shift_matrix_before.csv` preserves the independent reconstruction of the saved old rule.

The independent calculation uses the four workers' frozen source records from the earlier saved workbook. All 16 records previously labelled `Leave/other exclusion` in this sample were checked against the saved source reasons and match the revised recognised-leave list. There are no AVAIL records among these four workers. The 7.6-hour allowance was read from the saved Settings table. The new saved refresh matches these projections on every worker/date/shift row.

| Worker ID | Shifts | Refreshed positive `ResDayShift` rows | Pre-repair saved positive rows |
| --- | ---: | ---: | ---: |
| 19822 | 84 | 25 | 0 |
| 19835 | 84 | 5 | 0 |
| 24562 | 84 | 32 | 0 |
| 25376 | 84 | 46 | 0 |

The current workbooks expose `FilePathUrl` as a named single cell. Both edited Power Query resolvers now normalize Excel's `Column1` named-cell output to the canonical `FilePath` column, while also accepting a one-row `FilePath` table. No new workbook path table is required for these existing cells. The B `.m` and embedded workbook definition also differ; following the user's source precedence, the `.m` was edited and the workbook was not used to overwrite it.

## Saved result before the source repair

| Worker ID | AIN shifts checked | Saved `WorkerShiftSegments` hours | Saved `ResDayShift` rows |
| --- | ---: | ---: | ---: |
| 19822 | 84 | 0 on all 84 | 0 |
| 19835 | 84 | 0 on all 84 | 0 |
| 24562 | 84 | 0 on all 84 | 0 |
| 25376 | 84 | 0 on all 84 | 0 |

The roster covers 20 July through 16 August 2026, with AM, PM, and NIGHT shifts each day. All four workers pass `ReconciledWorkers_Eligible` and appear in `Availability-StaffList`. None has a valid `AVAIL` record in this filtered extraction. The **saved workbook's older query** starts each shift as available, then subtracts exclusions. All 336 saved shifts end at zero because that older query promotes each affected calendar day to a **whole-day exclusion** whenever combined excluded time reaches the Settings `ShiftDuration` of 7.6 hours. The independent reconstruction in `build_matrix.py` matches all 336 saved `WorkerShiftSegments` hours and the empty saved `ResDayShift` table.

## Example: 20 July 2026

| Worker ID | Source records on the date | Distinct excluded hours | Current rule | Example result |
| --- | --- | ---: | --- | --- |
| 19822 | `UNAVAIL` 06:00–22:00, source record 3600 | 16 | Exclude 00:00–24:00 | AM, PM, NIGHT: zero |
| 19835 | `UNAVAIL` 00:00–24:00, source record 3708; PEL 14:00–22:00, record 3709 | 24, with overlap counted once | Exclude 00:00–24:00 | AM, PM, NIGHT: zero |
| 24562 | `UNAVAIL` 00:00–14:00 and 22:00–24:00, records 1943–1944 | 16 | Exclude 00:00–24:00 | PM 14:00–22:00: zero despite no raw `UNAVAIL` overlap |
| 25376 | `UNAVAIL` 00:00–06:00 and 22:00–24:00, records 1534–1535 | 8 | Exclude 00:00–24:00 | AM and PM: zero despite no raw `UNAVAIL` overlap |

NIGHT shifts cross midnight. The matrix lists evidence and the daily exclusion decision for each calendar date touched by the shift. `ExclusionsAppliedToShift` shows the *expanded* exclusion clipped to that shift; `RemainingAvailableSegments` and `CalculatedAvailableHours` then show the subtraction. `ResDayShift` publishes only positive `EffectiveShiftHrs` rows, so zero-hour worker-shifts disappear there.

## Where Capacity enters

`Availability-StaffList` lists eligible workers independently of their available hours, which explains why all four can appear in a staff list. The AIN distribution path then needs a positive `ResDayShift` row for original availability: `CapacityDistrib(A.1)-shifts` imports the saved `ResDayShift` table, joins the worker name to a Resource and the role/date/shift to a Period, and assigns one unit of availability per matched row. A.2 caps those rows, B calculates final C### availability, and `Capacity.xlsx` imports B.

The saved workbooks are out of sequence. Treat this as a historical mismatch to investigate after an ordered rerun, not a verdict on the revised calculation:

| Saved workbook | Last saved | Four-worker evidence |
| --- | --- | --- |
| AIN `CapacityDistrib(A.2)-shifts.xlsx` | 18 Sep, 09:39 | Positive C# rows: Resource 10 = 27, 12 = 24, 26 = 25, 27 = 5 |
| AIN `CapacityDistrib(A.1)-shifts.xlsx` | 18 Sep, 14:03 | No imported `ResDayShift` rows or `ResPeriodAvailabilityTABLE` rows for the four workers |
| AIN `CapacityDistrib(B)-shifts.xlsx` | 19 Sep, 10:53 | Positive C### rows: Resource 10 = 15, 12 = 18, 26 = 18, 27 = 5; these rows have blank `ResAvailability`/`ResMaxAvail` |
| `Capacity.xlsx` | 19 Sep, 11:12 | Imports those same positive B C### counts |
| `Capacity-ShiftAvailability.xlsx` | 19 Sep, 15:22 | All four workers have zero hours and no `ResDayShift` rows |

The Resource mapping in the saved A.1 skeleton is 25376 → 10, 24562 → 12, 19822 → 26, and 19835 → 27. The A.2 file was last saved before the A.1 file; their saved tables disagree for all four workers. A.2's query sets `AvailabilityCapped` to zero when original Availability is null or the period is marked `UnavailableAllocatedPeriod`. B has positive C# values from A.2 while its original-availability join against A.1 leaves `ResAvailability` and `ResMaxAvail` blank for these four. That join did not gate the saved C### calculation, so Capacity imported positive B rows. Those pre-repair outputs formed a mixed-version chain, with no evidence of a deliberate availability override. Last-saved times alone do not prove the exact refresh order. Availability has since been refreshed and reconciled; the downstream chain has not.

Availability is now synchronized, refreshed and reconciled for the four workers. To judge downstream Capacity, synchronize the authoritative AIN B `.m` under its separate workbook authorization, then refresh StaffListMaster, AIN A.1, A.2 and B in order, plus runner-declared prerequisites. A.1 reads both Availability and StaffListMaster. Refresh Capacity afterwards if its published output is being judged. Compare the four IDs and source versions at every stage. The ordered rerun will establish whether the former mixed-version Capacity discrepancy is resolved.

The 18 September 09:30 full batch (`20260918-093026-9249c1e7`) completed 60/60 jobs, with Unit1 Availability, AIN A.1, A.2, B, and Capacity saved in sequence by 09:41. Its input snapshots show every downstream stage used the *earlier* Availability workbook saved by that batch. Later workbook saves produced today's mixed state: A.1 changed at 14:03 on 18 September, B and Capacity changed on 19 September before today's local Availability save, and A.2 still matches its 18 September 09:39 batch output. The 19 September batch folders contain organisation jobs only; they do not establish a later same-run Unit1 sequence. Successful completion of the 18 September run therefore does not validate the current four-worker filter against the current Capacity table.

## How to read the before matrix

1. Open `day_shift_matrix_before.csv`. Start with `ShiftStart` and `ShiftEnd` and the three source-record columns. Record numbers identify rows in the saved `IMPORT Availability Leave Sourc` table.
2. Read `DistinctDailyExcludedHoursAndRule`. The current query unions overlapping exclusion intervals on each calendar day and changes a total of at least 7.6 hours into a whole-day exclusion.
3. Compare `BaselineWithinShift`, `ExclusionsAppliedToShift`, and `RemainingAvailableSegments`. These explain `CalculatedAvailableHours`.
4. Compare that value with `SavedWorkerShiftSegmentsHours`. `FullShift` and the 7.6-hour allowance produce `CalculatedEffectiveShiftHrs`.
5. `PresentInSavedResDayShift` confirms whether the row survived the positive-hours publication filter.

## Source state and limits

- The saved workbook and `day_shift_matrix_before.csv` use the old false/true classification and promote all exclusions to whole days. The authoritative `.m` has now been edited to separate AVAIL, UNAVAIL and recognised leave, with a leave-only daily threshold and validation checks.
- A revised implementation with `AvailabilityRecordKind` and `AvailabilityDailyWindows` was embedded in the workbook on 13 September and was still embedded on 17 September. The adjacent `.m` source reverted to older logic in a 14 September commit (`19c4846`); the workbook reverted by the 19 September commit (`96f42f4`). Before this repair, the workbook and adjacent `.m` matched on the old rule. The 19 September 16:32 refreshed workbook now embeds the authoritative revised `.m`; read-only extraction confirmed an exact match.
- `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AvailabilityRecords_Explained.md` describes the revised rule. `docs/DATExx-Whiddon-Capacity-Shift-Availability-Rules.md` has been corrected to match it.
- The four-worker filters were retained through the bounded review, including the corrected Worker's Report ID, and were then removed from the `.m` source. Full-population workbook results remain pending source synchronization and refresh.
- The saved downstream tables do not yet constitute a same-run reconciliation to this four-worker sample. Availability alone has been synchronized and refreshed.

`build_matrix.py` regenerates the old-rule saved analysis as `day_shift_matrix_saved.csv` when Python and `openpyxl` are available. `build_expected_matrix.ps1` uses the frozen before CSV to regenerate the revised-rule expected matrix without opening a workbook. It writes both `day_shift_matrix_expected.csv` and the review-facing `day_shift_matrix.csv`.
