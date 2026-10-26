# Actual run: 20261026-145546-dca25314

Status: PartialOrFailed. Elapsed: 00:08:15. Peak recorded concurrency: 3.
Union of observed active intervals: 493.371 seconds. Gaps including resume pauses: 2 seconds.

Technical refresh only; business reconciliation not evaluated.
Comparison baseline: original illustrative 17m 23s. Difference does not establish performance improvement.

**OUTPUTS NOT VERIFIED/UPDATED: 45 of 55.**

Metadata assessment only; later manual saves are not runner receipts. See output-status.md for every output.
- TE/DemandInput: NotUpdated; Not updated by this run (Failed). Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/Availability: NotUpdated; Not updated by this run (Failed). Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK. The output predates a current declared input.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
- TE/Staff: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Availability did not complete. The output predates a current declared input.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
- TE/Intervals: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/DemandInput did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/DemandIntervals: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Intervals did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/Demand: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/DemandIntervals did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/Shifts: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Intervals did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/ShiftAverage: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Shifts did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/Allocation: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/ShiftAverage did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/Distribution: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Allocation did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/Role:AIN/1: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:AIN/2: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Staff did not complete.
- TE/Role:AIN/3: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Staff did not complete.
- TE/Role:RN/1: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:RN/2: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Staff did not complete.
- TE/Role:RN/3: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Staff did not complete.
- TE/Role:EN/1: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:EN/2: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Staff did not complete.
- TE/Role:EN/3: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Staff did not complete.
- TE/Role:CS/1: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:CS/2: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Staff did not complete.
- TE/Role:CS/3: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Staff did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Staff did not complete.
- TE/Capacity: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Role:AIN/3 did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Staff did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Staff did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Staff did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Staff did not complete.
- TE/Effort: NotUpdated; Not updated by this run (Blocked). Prerequisite TE/Capacity did not complete. The output predates a current declared input.
  Root cause: TE/DemandInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
  Root cause: TE/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Staff did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Staff did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Staff did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Staff did not complete.
  Root cause: TE/Allocation: Prerequisite TE/ShiftAverage did not complete.
  Root cause: TE/Distribution: Prerequisite TE/Allocation did not complete.
- JH-RY/Workers: NotUpdated; Not updated by this run (Failed). Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
- JH-RY/Availability: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Workers did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
- JH-RY/Staff: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Availability did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
- JH-RY/Shifts: NotUpdated; Not updated by this run (Failed). Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class The output predates a current declared input.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/ShiftAverage: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Shifts did not complete. The output predates a current declared input.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/Allocation: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/ShiftAverage did not complete. The output predates a current declared input.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/Distribution: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Allocation did not complete. The output predates a current declared input.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/Role:AIN/1: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/Role:AIN/2: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:AIN/3: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:RN/1: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/Role:RN/2: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:RN/3: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:EN/1: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/Role:EN/2: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:EN/3: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:CS/1: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
- JH-RY/Role:CS/2: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:CS/3: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Staff did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Capacity: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Role:AIN/3 did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Staff did not complete.
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Staff did not complete.
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Staff did not complete.
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Staff did not complete.
- JH-RY/Effort: NotUpdated; Not updated by this run (Blocked). Prerequisite JH-RY/Capacity did not complete. The output predates a current declared input.
  Root cause: JH-RY/Workers: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
  Root cause: JH-RY/Shifts: Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Staff did not complete.
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Staff did not complete.
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Staff did not complete.
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Staff did not complete.
  Root cause: JH-RY/Allocation: Prerequisite JH-RY/ShiftAverage did not complete.
  Root cause: JH-RY/Distribution: Prerequisite JH-RY/Allocation did not complete.

## Failed or stopped attempts

- TE/DemandInput: Failed; Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ShiftUnitDemandHRS/ShiftUnitDemandHRS' failed foreground refresh: [Demand Extraction validation failed] Unique timing and exact shift spans: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Complete demand and fractional attendance conversion: Each enabled Role Group calendar cell requires exactly one settings shift definition.
Retained hours across all full and partial planning fortnights: Each enabled Role Group calendar cell requires exactly one settings shift definition.
- TE/Availability: Failed; Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
- JH-RY/Workers: Failed; Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'IMPORT Leave Balances/IMPORT_Leave_Balances' failed foreground refresh: [Expression.Error] The column 'Employee Code' of the table wasn't found.
- JH-RY/Shifts: Failed; Worker exit 1; phase Calculating; save confirmed: False; Unable to get the Calculate property of the Application class

## Untimed or incomplete work

- TE/Staff: Blocked; Prerequisite TE/Availability did not complete.
- TE/Intervals: Blocked; Prerequisite TE/DemandInput did not complete.
- TE/DemandIntervals: Blocked; Prerequisite TE/Intervals did not complete.
- TE/Demand: Blocked; Prerequisite TE/DemandIntervals did not complete.
- TE/Shifts: Blocked; Prerequisite TE/Intervals did not complete.
- TE/ShiftAverage: Blocked; Prerequisite TE/Shifts did not complete.
- TE/Allocation: Blocked; Prerequisite TE/ShiftAverage did not complete.
- TE/Distribution: Blocked; Prerequisite TE/Allocation did not complete.
- TE/Role:AIN/1: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:AIN/2: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:AIN/3: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:RN/1: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:RN/2: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:RN/3: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:EN/1: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:EN/2: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:EN/3: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:CS/1: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:CS/2: Blocked; Prerequisite TE/Staff did not complete.
- TE/Role:CS/3: Blocked; Prerequisite TE/Staff did not complete.
- TE/Capacity: Blocked; Prerequisite TE/Role:AIN/3 did not complete.
- TE/Effort: Blocked; Prerequisite TE/Capacity did not complete.
- JH-RY/Availability: Blocked; Prerequisite JH-RY/Workers did not complete.
- JH-RY/Staff: Blocked; Prerequisite JH-RY/Availability did not complete.
- JH-RY/ShiftAverage: Blocked; Prerequisite JH-RY/Shifts did not complete.
- JH-RY/Allocation: Blocked; Prerequisite JH-RY/ShiftAverage did not complete.
- JH-RY/Distribution: Blocked; Prerequisite JH-RY/Allocation did not complete.
- JH-RY/Role:AIN/1: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:AIN/2: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:AIN/3: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:RN/1: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:RN/2: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:RN/3: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:EN/1: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:EN/2: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:EN/3: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:CS/1: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:CS/2: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Role:CS/3: Blocked; Prerequisite JH-RY/Staff did not complete.
- JH-RY/Capacity: Blocked; Prerequisite JH-RY/Role:AIN/3 did not complete.
- JH-RY/Effort: Blocked; Prerequisite JH-RY/Capacity did not complete.
