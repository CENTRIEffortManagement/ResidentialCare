# ShiftPeriod refresh selection

## Active run update

The user excluded `Capacity-ShiftAvailability.xlsx` and `StaffListMaster.xlsx` and instructed refresh of the remaining listed files. Active selection: Unit1/global ranges **2-6 and 9-24**, 21 workbooks total, using the listed saved roles AIN, AINC4 and RN. The 23-workbook table below records the original selection; positions 7 and 8 are excluded from this run.

Both exact ranges passed `-ValidateSelectionOnly` with exit code 0 on 2026-09-10. Range 2-6 started at 18:28:28 in hidden Excel with the default 30-minute per-workbook timeout. Range 9-24 must start only after successful completion of range 2-6.

First range log: `CLIENT/DATExx-Whiddon/UNITS/RunLogs/AllUnitsRefresh-20260910-182829.log`.

Scope: Unit1 sequence 2-24 inclusive; 23 workbooks. With only Unit1 discovered, these are also global sequence 2-24.

Entry point: CLIENT/DATExx-Whiddon/Run-RefreshRunner.cmd

The corrected ShiftPeriod table is assumed saved in 2. Calculations/Settings Data.xlsx, as supplied by the user. Demand extraction reads that saved table before Settings Data refreshes its own queries. No M synchronization is included.

| Global / Unit1 sequence | Workbook relative to Unit1 |
| --- | --- |
| 2 | 1. Input/2-DemandExtract.xlsx |
| 3 | 2. Calculations/Settings Data.xlsx |
| 4 | 2. Calculations/Intervals.xlsx |
| 5 | 2. Calculations/DemandIntervals.xlsx |
| 6 | 2. Calculations/Demand.xlsx |
| 7 | 2. Calculations/Capacity-ShiftAvailability.xlsx |
| 8 | 2. Calculations/StaffListMaster.xlsx |
| 9 | 2. Calculations/Shifts.xlsx |
| 10 | 2. Calculations/AllocationByShiftAverage.xlsx |
| 11 | 2. Calculations/Allocation.xlsx |
| 12 | 2. Calculations/Shift-StaffDistribution.xlsx |
| 13 | 2. Calculations/MutliRoleCheck.xlsx |
| 14 | 2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx |
| 15 | 2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx |
| 16 | 2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx |
| 17 | 2. Calculations/AINC4/CapacityDistrib(A.1)-shifts.xlsx |
| 18 | 2. Calculations/AINC4/CapacityDistrib(A.2)-shifts.xlsx |
| 19 | 2. Calculations/AINC4/CapacityDistrib(B)-shifts.xlsx |
| 20 | 2. Calculations/RN/CapacityDistrib(A.1)-shifts.xlsx |
| 21 | 2. Calculations/RN/CapacityDistrib(A.2)-shifts.xlsx |
| 22 | 2. Calculations/RN/CapacityDistrib(B)-shifts.xlsx |
| 23 | 2. Calculations/Capacity.xlsx |
| 24 | 2. Calculations/Effort.xlsx |

Preflight: all 23 files exist, no selected Excel temporary lock files found, pwsh available, enabled saved roles AIN/AINC4/RN match the manifest in order. No active refresh process was found; one existing Excel process was present and has not been closed.

The calculation and Unit1 shortcuts reference absent bin/runner scripts. The intact approved top-level runner supports this exact inclusive range. No runner files have been modified.

Formal selection validation and refresh are pending the role-list choice required by docs/ResidentialCare-Runner-Agent-Instructions.md. No workbook has been opened, refreshed or saved by this task.

The user confirms the night crossover has already been handled in M. This task is limited to propagating the corrected ShiftPeriod settings.

## Final run outcome: stopped by user

Stopped at 2026-09-10 18:41:33 during Intervals.xlsx. The approved stop-now command was issued; Excel was blocked in connection polling, so the exact automation Excel process created for Intervals.xlsx (PID 18200, started 18:29:58) was identity-checked and terminated. The runner then observed the stop request and exited.

Completed, saved and closed in hidden Excel: 1. Input/2-DemandExtract.xlsx and 2. Calculations/Settings Data.xlsx. Intervals.xlsx was not saved by this refresh; its unsaved refresh was discarded. No later workbook started, including range 9-24. The two user-excluded workbooks were not refreshed.

Validation: both requested ranges passed, exit code 0. Run termination: unit stage exit code 3 / STOPPED; top-level wrapper exit code 1 / FAILED because it maps the child stop code to failure. Other Excel instances were not terminated.

Final status: CLIENT/DATExx-Whiddon/UNITS/RunLogs/current-status.txt. Log: CLIENT/DATExx-Whiddon/UNITS/RunLogs/AllUnitsRefresh-20260910-182829.log.

## Resumed run: stopped by user

After the user reported Intervals complete, ranges 5-6 and 9-24 passed validation (exit 0). Range 5-6 began in hidden Excel at 2026-09-10 18:47:36. DemandIntervals.xlsx stalled in connection monitoring. The user requested stop; the stop-now command was issued and the identity-verified automation Excel instance (PID 50316, created 18:47:37) was terminated to release the blocked call.

No workbook completed or was saved in this resumed run. DemandIntervals.xlsx unsaved refresh was discarded; Demand.xlsx and range 9-24 never started. Unit stage reports STOPPED (exit 3); top-level wrapper exits 1. No M synchronization or load inspection was performed.

Log: CLIENT/DATExx-Whiddon/UNITS/RunLogs/AllUnitsRefresh-20260910-184736.log.

## Isolated retry: succeeded

At the user request, retried only Unit1/global sequence 5, DemandIntervals.xlsx. Selection validation passed, exit 0. Before refresh, memory use was 43.9% with 17.91 GB available and no existing Excel process. Hidden Excel opened the workbook at 18:57:44, started refresh at 18:57:49, successfully polled 14 connections, saved and closed at 18:58:23, and quit. Top-level runner succeeded at 18:58:24 with exit code 0.

This successful retry supports reduced memory pressure or a cleaner Excel session as possible contributing factors; it does not isolate the cause of earlier freezes or measure peak memory. No M code or runner code was changed. Remaining workbooks were not included in this isolated test.

Log: CLIENT/DATExx-Whiddon/UNITS/RunLogs/AllUnitsRefresh-20260910-185741.log.

## Final continuation and revised exclusions

The user authorized all remaining files, with no memory checks. Range 6-6 and 9-24 passed validation. Demand.xlsx refreshed, saved and closed successfully (exit 0; log AllUnitsRefresh-20260910-190043.log). Range 9-24 then ran sequentially in hidden Excel.

Completed and saved before the revised capacity exclusions took effect: Shifts.xlsx, AllocationByShiftAverage.xlsx, Allocation.xlsx, Shift-StaffDistribution.xlsx, MutliRoleCheck.xlsx, all three AIN capacity-distribution workbooks and all three AINC4 capacity-distribution workbooks.

The user then excluded all role capacity-distribution workbooks and Capacity.xlsx. Stop-now was issued immediately. The first RN capacity workbook had already started; it was stopped and closed without saving. The other two RN workbooks and Capacity.xlsx did not run. Unit exit 3 / top-level exit 1 reflects the requested stop. No completed workbook changes were rolled back.

Only sequence 24, Effort.xlsx, was then validated and refreshed, saved and closed successfully at 2026-09-10 19:13:59 (exit 0). It uses the existing saved Capacity.xlsx results, as that workbook was excluded. No automation workbook was left open. No M synchronization, load inspection or memory checks were performed during this continuation.

Logs under CLIENT/DATExx-Whiddon/UNITS/RunLogs/: AllUnitsRefresh-20260910-190043.log (Demand), AllUnitsRefresh-20260910-190207.log (sequence 9 onward; verify filename against latest run records), AllUnitsRefresh-20260910-191308.log (Effort).
