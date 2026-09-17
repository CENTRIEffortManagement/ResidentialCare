# Runner benchmark — 17 September 2026

These measurements concern technical refresh execution, not business reconciliation.
Historical logs identify dates, not hardware. The pre-15 September cutoff is the
operator's proxy for the previous PC. Data, Excel versions and cache conditions
were not controlled across dates; these are not CPU-only benchmarks.

## Expanded runner live verification

Unit1 run `20260917-170909-5106fe61` completed all 26 selected files in approximately
13 minutes 30 seconds, exit 0. All 26 supervisor receipts confirmed save/close
completion and owned-process exit. Output metadata matched the completion records.
Peak concurrency was four batches; the configured maximum was seven.
Demand-MasterRoster recovered from a brief missing identity and required owned
process cleanup after confirmed save. No Git-reader encounter was recorded, so
this run did not exercise the new Git allowance under observed contention.

The matching 14-file section (legacy Unit1 IDs 11–24) completed in 3m 49s versus
11m 04s in `UNITS/RunLogs/AllUnitsRefresh-20260911-145025.log`: approximately 65%
less elapsed time, or 2.9 times the throughput. The nine C2 files took 2m 23s
versus 7m 25s. This does not establish seven-batch production capacity; seven was
tested with synthetic workers, while this live run reached four.

## Same-day serial comparison

Run `20260917-174639-0b4a3552` selected only the A.1 file in each enabled Unit1
role, with maximum concurrency one. All three completed, exit 0. Declared input
size/timestamp metadata matched the preceding parallel run.

| Role | 11 September sequential | Current parallel | Current serial |
|---|---:|---:|---:|
| AIN | 57s | 49.86s | 48.56s |
| AINC4 | 56s | 52.06s | 48.24s |
| RN | 58s | 52.27s | 48.06s |

Serial execution reduced average per-file duration by approximately 6% compared
with the same-day parallel run. Current and historical refresh-call durations
are not directly comparable: the current worker temporarily disables supported
background refresh and uses different readiness/cleanup checks.

## Archived-worker comparison

The operator explicitly approved an isolated test using the unchanged worker
from commit `47e11a2` (8 September), source
`CLIENT/DATExx-Whiddon/UNITS/runner/src/Invoke-ExcelWorkbookRefresh.ps1`.
Its Git blob was verified as `de79724af2ba18f63e6f266a5d0643c47905331f`.
The installed runner was not replaced. Tests used hidden Excel, the original
polling behaviour, a shared exclusion lock and an external 30-minute worker
deadline. The archived worker has older save/timeout protections; its success
messages are not equivalent to modern supervisor receipts.

| Unit1 workbook | Historical baseline | Refresh call old → current | Total logged runtime old → current |
|---|---|---:|---:|
| AIN/CapacityDistrib(A.1)-shifts.xlsx | 11 September, 14:50 run | 7s → 3s | 57s → 38s |
| Intervals.xlsx | 10 September, 15:37 run | 4s → 2s | 43s → 39s |
| AllocationByShiftAverage.xlsx | 10 September, 19:02 run | 4s → 2s | 71s → 53s |
| Capacity-ShiftAvailability.xlsx | 12 September, 12:40 run | 4s → 2s | 54s → 45s |

Historical baseline files are under `CLIENT/DATExx-Whiddon/UNITS/RunLogs/`:
`AllUnitsRefresh-20260911-145025.log`, `AllUnitsRefresh-20260910-153734.log`,
`AllUnitsRefresh-20260910-190207.log`, and `AllUnitsRefresh-20260912-124009.log`.
Current local records are under `RunLogs/ArchivedBenchmarks/`, in
`20260917-AIN-A1-47e11a2` and `20260917-180758-three-calculations`.

All four archived tests returned exit 0, logged save/close success, changed output
metadata and left no Excel process running at final checks. Only selected files
were refreshed; upstream and downstream files were not additionally refreshed.
An initial Tableau hold on AllocationByShiftAverage blocked validation; the
operator cleared it before the successful rerun. No application was killed to
bypass that hold.

Refresh calls took approximately 50–57% less time, but may return before background
work finishes. Total workbook runtimes were approximately 9–33% shorter. Log
timestamps resolve only to whole seconds; decimal whole-process timings must not
be presented as decimal refresh-call measurements. Fifty percent less time means
twice as fast, not fifty percent faster.

More Units provide additional independent ready work and may better utilise the
seven-batch ceiling. Actual throughput and memory capacity must be measured;
neither sevenfold speedup nor memory pressure as the cause of historical crashes
has been established.
