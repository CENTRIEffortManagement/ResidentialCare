# ParallelInputs-Unit1Priority implementation and trial checkpoint

Implemented 18 September 2026. Current remains the default sequence.
The separate profile starts A1/S1/D1 independently, prioritises ready Unit1
work and fills spare slots from Unit2. The existing coordinator remains the
resource manager, with a maximum of seven active workers.

## Review files

- [Profile instructions](ResidentialCare-Runner-ParallelInputs-Unit1Priority.md)
- [Unit1 planned Gantt](ParallelInputs-Unit1Priority-preview-20260918-090128-cb10488a-planned.mmd)
- [Unit1 timing provenance](ParallelInputs-Unit1Priority-preview-20260918-090128-cb10488a-forecast.json)
- [All-Units planned Gantt](ParallelInputs-Unit1Priority-preview-20260918-090130-c0b1f608-planned.mmd)
- [All-Units unknown timing notes](ParallelInputs-Unit1Priority-preview-20260918-090130-c0b1f608-planned-notes.md)
- [Original illustrative 17m 23s baseline](batch-duration-gantt.mmd)

Unit1's historical-duration forecast is 652.4 seconds (approximately 10m 52s),
with peak concurrency four. The all-Units forecast reaches seven active workers
and stops its timed prefix at Org/History because no completed measurement is
available. Its end-to-end duration remains unknown; Cost also lacks a measurement.
These forecasts are simulations, not completed live runs.

The original baseline is unchanged, verified by SHA256:
F29199C84B672FD8394203D73CB650E9358A3D09186AA1706256F735FE659DCD.
Its old total and concurrency assumptions remain unverified.

## Verification

- New profile, priority, forecast, durable attempts, resume and relocation:
  1,467 checks passed, including actual retry-chart rendering.
- Existing batch suite: 149 assertions passed.
- Recovery suite: 55 checks passed.
- Launcher suite: 68 checks passed.
- File-access/identity suite: 26 checks passed.
- Git-aware suite: 52 checks passed; optional live-process checks not run.
- Excel safety/supervisor mock cases passed.
- Seven-role mock harness: all 21 synthetic jobs passed, with no Excel startup.
- Legacy status, RN three-file validation and global sequence 1 validation passed.
- Unit1 forecast (26 bars) and all-Units forecast (56 timed bars) rendered using
  the locally installed Mermaid renderer with headless Chromium.
- Diff whitespace validation passed; Git reported only line-ending conversion warnings.

The seven-role harness now builds isolated fixtures rather than depending on
missing production TestRole folders. The Unit1 RN legacy smoke entry point had
stale launcher, manifest and worker-folder paths; these were repaired using
existing scripts and relative paths.

The named Mermaid extension tools were unavailable in this session. Rendering
was verified with its installed Mermaid bundle instead; no extension UI preview
or cloud sync was invoked.

## Initial validation block (resolved for Unit1)

Scope: Unit1 only, 26 files in 11 batches, saved AIN/AINC4/RN roles,
ParallelInputs-Unit1Priority, maximum seven workers.

Validation returned exit 1 because Tableau was using:

- Unit1 Shifts.xlsx — Tableau PID 46696.
- Unit1 AllocationByShiftAverage.xlsx — Tableau PID 61712.
- Unit1 Effort.xlsx — Tableau PID 51808.

All three are beneath CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/.
Windows Restart Manager checks required execution outside the sandbox; the
same checks succeeded there and identified these external file users.
No guard was bypassed and no external application was stopped.

That initial trial did not start: zero workbooks refreshed or saved, and no workbook
opened or left open by this task. Existing Excel processes were left untouched.
At that checkpoint there was no live-run state or actual Gantt. The validation findings
are recorded here; the planned snapshots are in RunLogs/BatchRefresh beneath
the configured client root, with their preview IDs matching the review files.

## Unit1 trial completed

After the operator closed Tableau, Unit1 validation passed and the authorized
hidden-Excel trial completed on 18 September 2026, using the saved roles.
Run: `20260918-091108-616ed35f`; exit 0; all 26 files completed with verified
saves and closes, with no inspection flags, recovery snapshots or report errors.
Measured elapsed time: 679.389 seconds (11m 19s); peak concurrency: four.
A1, S1 and D1 were admitted within the first second. The prior Unit1 forecast
was 652.4 seconds; actual elapsed was approximately 27 seconds longer.
This Unit1-only result is not a like-for-like comparison with the original
illustrative full-sequence 17m 23s baseline, which remains unchanged.

- [Unit1 actual Gantt](ParallelInputs-Unit1Priority-20260918-091108-616ed35f-actual.mmd)
- [Recorded timings](ParallelInputs-Unit1Priority-20260918-091108-616ed35f-actual.json)
- [Actual run notes](ParallelInputs-Unit1Priority-20260918-091108-616ed35f-actual-notes.md)
- [Frozen run forecast](ParallelInputs-Unit1Priority-20260918-091108-616ed35f-planned.mmd)

All 26 chart intervals agree with the recorded start/finish timestamps;
every selected predecessor finished before its successor started. The actual
Mermaid chart rendered successfully with 26 bars. Run records and coordinator
log are under `CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh/20260918-091108-616ed35f/`.

## Full rollout initial validation block (resolved)

The subsequent all-Units plus organisation validation selected 60 files in
27 batches, but returned exit 1. Windows identified Tableau background process
PID 46340 (`tabprotosrv.exe`, parent `tableau.exe`) using
`CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All.xlsx`.
No Excel was started for the full run, and no full-run workbook was refreshed
or saved. No external process was stopped and no guard was bypassed.

The operator subsequently requested continuation. Revalidation passed and the
fresh full run below included Unit1 again.

## Full run completed

Run `20260918-093026-9249c1e7` completed on 18 September 2026, exit 0:
all 60 workbooks in 27 batches (26 per Unit and eight organisation workbooks)
refreshed in hidden Excel with verified saves and closes. Saved roles remained
AIN, AINC4 and RN. There are no incomplete jobs, inspection flags, recovery
snapshots or report errors. Business reconciliation remains unevaluated.

Measured elapsed time was approximately **15m 21s**, with peak recorded
concurrency **seven**. All 60 chart intervals agree with the recorded timestamps;
every predecessor finished before its successor started. Both actual (60 bars)
and frozen planned (56 timed bars) Mermaid files rendered successfully.

- [Full actual Gantt](ParallelInputs-Unit1Priority-20260918-093026-9249c1e7-actual.mmd)
- [Actual timing evidence](ParallelInputs-Unit1Priority-20260918-093026-9249c1e7-actual.json)
- [Actual run notes](ParallelInputs-Unit1Priority-20260918-093026-9249c1e7-actual-notes.md)
- [Frozen planned Gantt](ParallelInputs-Unit1Priority-20260918-093026-9249c1e7-planned.mmd)
- [Forecast provenance](ParallelInputs-Unit1Priority-20260918-093026-9249c1e7-forecast.json)

The frozen forecast still has no full-duration estimate because History and Cost
lacked successful prior observations. This run supplies those observations for
future forecasts; the pre-run forecast has not been retroactively filled in.
The observed full run is roughly two minutes shorter than the original
illustrative 17m 23s, but that baseline had unverified timings and exceeded
seven concurrent jobs. This difference is not proof of a speed improvement or
of a globally optimal schedule. Current remains the default fallback profile.

Observed scheduling evidence:

- Unit1 A1/S1/D1 started first, then Unit2's ready input jobs used spare slots.
- Unit2 Intervals released Shifts while its DemandIntervals was still running.
- Organisation StaffAll started after both Units' Staff outputs, before either
  Unit's full sequence had finished.
- At 09:38:17, newly ready Unit1 RN took the vacancy ahead of waiting Unit2
  MultiRole work; no running Unit2 job was interrupted.
- Organisation EffortAll waited for both Units' Effort outputs. History and
  Inefficiencies then overlapped, followed by HistoryRead and Cost.

Run state, worker receipts and coordinator log are preserved under
`CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh/20260918-093026-9249c1e7/`.

Workbook imports have not been independently inspected. The static startup-input
relationship follows the operator's confirmation; declared input guards remain
active. Technical refresh completion does not certify business reconciliation.
