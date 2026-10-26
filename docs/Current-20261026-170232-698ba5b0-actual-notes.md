# Actual run: 20261026-170232-698ba5b0

Status: PartialOrFailed. Elapsed: 00:19:38. Peak recorded concurrency: 12.
Union of observed active intervals: 1174.759 seconds. Gaps including resume pauses: 3 seconds.

Technical refresh only; business reconciliation not evaluated.
Comparison baseline: original illustrative 17m 23s. Difference does not establish performance improvement.

**OUTPUTS NOT VERIFIED/UPDATED: 7 of 90.**

Metadata assessment only; later manual saves are not runner receipts. See output-status.md for every output.
- Org/Cost: NotUpdated; Not updated by this run (Failed). Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'TargetRatebyDate/Table_TargetRatebyDate' failed foreground refresh: [DataSource.NotFound] File or Folder: Could not find a part of the path 'C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx'. The output predates a current declared input.
  Root cause: Org/Cost: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'TargetRatebyDate/Table_TargetRatebyDate' failed foreground refresh: [DataSource.NotFound] File or Folder: Could not find a part of the path 'C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx'.
- Org/History: NotUpdated; Not updated by this run (Failed). Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table. The output predates a current declared input.
  Root cause: Org/History: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table.
- Org/HistoryRead: NotUpdated; Not updated by this run (Blocked). Prerequisite Org/History did not complete.
  Root cause: Org/History: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table.
- Org/Reporting: NotUpdated; Not updated by this run (Blocked). Prerequisite Org/HistoryRead did not complete. The output predates a current declared input.
  Root cause: Org/Cost: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'TargetRatebyDate/Table_TargetRatebyDate' failed foreground refresh: [DataSource.NotFound] File or Folder: Could not find a part of the path 'C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx'.
  Root cause: Org/History: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table.
- Org/LeaveBalanceStats: NotUpdated; Not updated by this run (Blocked). Prerequisite Org/Reporting did not complete.
  Root cause: Org/Cost: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'TargetRatebyDate/Table_TargetRatebyDate' failed foreground refresh: [DataSource.NotFound] File or Folder: Could not find a part of the path 'C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx'.
  Root cause: Org/History: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table.
- Org/LeaveBalanceReduction: NotUpdated; Not updated by this run (Blocked). Prerequisite Org/LeaveBalanceStats did not complete.
  Root cause: Org/Cost: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'TargetRatebyDate/Table_TargetRatebyDate' failed foreground refresh: [DataSource.NotFound] File or Folder: Could not find a part of the path 'C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx'.
  Root cause: Org/History: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table.
- Org/LeaveBalanceReporting: NotUpdated; Not updated by this run (Blocked). Prerequisite Org/LeaveBalanceReduction did not complete.
  Root cause: Org/Cost: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'TargetRatebyDate/Table_TargetRatebyDate' failed foreground refresh: [DataSource.NotFound] File or Folder: Could not find a part of the path 'C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx'.
  Root cause: Org/History: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table.

## Failed or stopped attempts

- Org/Cost: Failed; Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'TargetRatebyDate/Table_TargetRatebyDate' failed foreground refresh: [DataSource.NotFound] File or Folder: Could not find a part of the path 'C:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\Unit1\2. Calculations\Settings Data.xlsx'.
- Org/History: Failed; Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'Effort+OutcomesXY/EffortOutcomesAG0' failed foreground refresh: [Expression.Error] The key didn't match any rows in the table.

## Untimed or incomplete work

- Org/HistoryRead: Blocked; Prerequisite Org/History did not complete.
- Org/Reporting: Blocked; Prerequisite Org/HistoryRead did not complete.
- Org/LeaveBalanceStats: Blocked; Prerequisite Org/Reporting did not complete.
- Org/LeaveBalanceReduction: Blocked; Prerequisite Org/LeaveBalanceStats did not complete.
- Org/LeaveBalanceReporting: Blocked; Prerequisite Org/LeaveBalanceReduction did not complete.
