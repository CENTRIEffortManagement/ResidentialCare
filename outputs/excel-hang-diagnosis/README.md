# Effort-All refresh hang: 11 September 2026

The user reported that an ordinary refresh produced the grey, unresponsive Excel window, recurring roughly six times. Diagnosis was read-only; no workbook was opened, parsed, saved, refreshed, or terminated by this investigation.

## Captured evidence

- At 15:43:33 Australia/Sydney, Excel PID 23440, titled `Effort-All - Excel`, was still not responding. It started at 15:37:57.
- Seven Microsoft.Mashup.Container.Loader processes belonged to PID 23440 and started at 15:38:09-10.
- An 8.4-second sample showed zero read/write transfer increments in Excel and all seven workers. Excel accumulated 0.125 CPU seconds; each worker accumulated between zero and 0.03125 CPU seconds. This supports a stalled refresh rather than sustained calculation or file reading during the sample.
- About 12.2 GiB of physical memory was free out of 31.9 GiB at the final capture. Office was x64, version 16.0.20326.20112. There is no evidence of physical-memory exhaustion in this snapshot.
- The main/initial Excel thread, TID 42488, was waiting. Windows Wait Chain Traversal reported it blocked, returned no downstream owner, and detected no cycle. This does not identify the responsible component or rule out waits unsupported by WCT.
- A separate Excel process, PID 36392, had no main window. Its involvement has not been established.
- Captured Application events include an Excel Application Hang (1002) and AppHangB1 report at 18:45 on 10 September, using the same Excel version.
- Power Query worker crash clusters occurred at 18:41, 18:45, and 18:54 on 10 September and 14:38 on 11 September. Messages report an unexpected exception in a Mashup background thread and FailFast termination. These precede the current filepath edit; they do not prove the cause of this current hang. Some worker failures could follow termination of their host.
- Foxit PDF Creator is registered with LoadBehavior 3. This is a candidate for add-in isolation, not evidence that Foxit caused the hang.

## Interpretation and next isolation

Confirmed: a refresh-triggered Excel hang, with very little process activity and a history of earlier Excel/Power Query failures. The precise blocked component and query remain unproven. No current crash event identified a faulty module; no process dump or native call stack was captured.

After recovery, a controlled repeat with third-party COM add-ins disabled can test add-in involvement. If the hang persists, refresh individual output queries to locate the trigger, then compare background-refresh settings. These are proposed diagnostic steps; none was performed or configured here. Repeated hangs in the same Office build also justify checking Office updates/repair after isolation.

Microsoft references:

- https://support.microsoft.com/en-us/excel/excel-not-responding-hangs-freezes-or-stops-working
- https://learn.microsoft.com/en-us/windows/win32/api/wct/nf-wct-getthreadwaitchain

`capture.json` holds the process, memory, thread-wait and event evidence. `capture.ps1` is the read-only capture script; pass a current Excel PID explicitly for another incident. Process IDs are valid only for the captured instance.
