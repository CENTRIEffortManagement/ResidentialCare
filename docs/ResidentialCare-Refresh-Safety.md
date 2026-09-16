# All-units refresh readiness and recovery

Applies to `UNITS/Run-AllUnitsRefresh.cmd` and the integrated unit stage of
`Run-RefreshRunner.cmd` under `CLIENT/DATExx-Whiddon`. The older per-unit stage
engines and organisation engine do not yet use this supervisor.

## What changed

The workbook sequence was already serial: the next workbook waited for the
previous worker to return. Connection refresh flags alone did not establish
that Excel was ready to save. A synchronous Excel call could also block the
worker's own stop and timeout checks.

Each workbook now runs in an Excel worker with a separate PowerShell supervisor.
The supervisor makes no Excel COM calls and enforces the configured timeout
(30 minutes by default) across startup, opening, refresh, calculation, save,
close and process exit. Its heartbeat identifies the last recorded phase every
15 seconds. `StopMode Now` allows five seconds for cooperative cancellation,
then terminates only the registered Excel instance and its observed Power Query
children. PID, process name and creation time must match before termination.
`AfterCurrent` continues to allow the current workbook to finish.

Before saving, the worker:

1. Confirms Excel opened the exact requested path and did not open it read-only.
2. Captures supported OLE DB, ODBC and query-table `BackgroundQuery` settings and
   temporarily disables background refresh before calling `RefreshAll`. Original
   settings are restored before saving. Excel documents that `RefreshAll` otherwise
   runs objects with `BackgroundQuery` enabled asynchronously.
3. Allows Power Query a stop-responsive 15-second settle interval, waits for a
   successful quiet connection/query-table check, requests calculation, then requires
   three consecutive quiet checks of pollable workbook connections,
   worksheet query tables and Excel's calculation state. A failed or unknown
   readiness read cannot count as completion.
   `CalculateUntilAsyncQueriesDone` is deliberately not used because Excel can spin
   indefinitely inside that API for some Power Query workbooks under automation.
4. Restores the captured refresh settings, checks read-only state again and saves to the original path.
5. Confirms `Workbook.Saved` and the unchanged target path, then closes and quits.

Validation and workbook startup also ask Windows Restart Manager which processes
are using each selected file. External users (for example Tableau's
`tabprotosrv`) block the run with their application name and PID. The worker
repeats this check immediately before saving, allowing only its own registered
Excel process. This checks a gap left by Excel's `~$` lock file and a successful
read/write-open test: another application can allow file access while preventing
Excel's replace/rename save operation. The check only lists file users and never
closes other applications. Close Tableau workbooks using these outputs before
refreshing; reopen/reload them after Excel finishes. A file-usage inspection
failure blocks validation rather than assuming the file is free. There remains
a race if another application opens a file after the last check; the supervisor
still bounds the save in that case.

The supervisor requires the completed worker state plus exit of the registered
Excel process and observed Power Query children before reporting success. It
allows 15 seconds for normal exit. If the worker exited successfully with save,
close and Quit confirmed, it terminates only lingering owned processes, then
checks that they exited before continuing. This recovery is logged as a warning;
an interrupted save cannot use it. It fails the sequence if cleanup remains
incomplete, and reports separately whether saving was confirmed. This gate does not validate
business outputs or guarantee that Excel/Power Query cannot hang; it makes the
wait visible, bounded and stoppable.

## Diagnosing a wait

### Planned Tableau confirmation gate

Before a future refresh starts, detect running Tableau desktop and data-source
processes, including `tableau` and `tabprotosrv`. List available Tableau workbook
window titles and process IDs. Do not claim that window titles identify every
connected Excel source; background data-source processes can remain without a
visible workbook window.

If any Tableau session is running, prompt the operator to confirm that those
sessions will not affect the selected refresh. Default to cancelling; validation
and status-only commands remain noninteractive and report the finding. An agent
or noninteractive refresh must obtain the same explicit confirmation, scoped to
the selected run, before supplying any future acknowledgement option. Do not
carry this acknowledgement automatically into a different run.

This confirmation is additional to the implemented file-user check. It cannot
override a detected handle on a selected workbook: that file must be released,
then validation repeated. Do not close Tableau or discard unsaved Tableau work
automatically. If no Tableau processes are running, proceed with normal
validation without the extra prompt.

This section records the requested plan; the Tableau-session confirmation prompt
has not yet been implemented. The file-user check described above is implemented.

### Current diagnostics

Logs and `current-status.txt` are under `UNITS/RunLogs`. Per-workbook
`refresh-worker-*.json` records retain the last phase, owned Excel identity and
save-confirmation flag. Exit 3 means operator stop; exit 124 means the supervisor
deadline expired. Other nonzero exits stop the sequence as failures.

A forced stop during saving leaves save completion unconfirmed. Preserve any
Excel temporary files; do not rename one over the original or automatically
substitute it. Verify the original target is closed before restarting from that
workbook. No automatic retry occurs, and a refresh never synchronizes M source.

For example, with Unit1 first in the discovered order, resume only its workbooks
9 through 24 from the repository root:

```powershell
.\CLIENT\DATExx-Whiddon\UNITS\Run-AllUnitsRefresh.cmd -StartAtSequence 9 -EndAtSequence 24
```

Check the printed selection before proceeding; adding or removing earlier unit
folders can change global numbering. Omitting the end sequence includes later units.

## Validation

`pwsh -NoProfile -File scripts/test-excel-refresh-safety.ps1` exercises readiness,
process identity, false success, timeout and stop handling with mocked Excel
objects and disposable PowerShell workers. It does not open Excel. Run the three
smoke checks in `ResidentialCare-Runner-Agent-Instructions.md` after changes,
then validate the exact production selection before an authorized refresh.

Microsoft references:

- [Pending query drain](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculateuntilasyncqueriesdone)
- [Calculation state](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationstate)
- [Read-only state](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.readonly)
