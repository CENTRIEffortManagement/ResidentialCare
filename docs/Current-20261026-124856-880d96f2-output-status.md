# Output status: 20261026-124856-880d96f2

Historical run status: PartialOrFailed. Checked UTC: 2026-10-26T01:52:07.7838569Z.

**OUTPUTS NOT VERIFIED/UPDATED: 52 of 55.**

Filesystem metadata only. This checker did not open, refresh, save or change a workbook. A later manual save is not a runner completion receipt. Business results were not validated.

- Date/DemandMaster: **NotUpdated**; 1. Input/Demand-MasterRoster Manual Read.xlsx; last saved UTC: 2026-10-26T01:33:46.5826654Z. Not updated by this run (Blocked). Prerequisite TE/Settings did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- TE/AllocationInput: **NotUpdated**; UNITS/TE/1. Input/1-AllocationExtracted.xlsx; last saved UTC: 2026-10-26T01:29:47.2155078Z. Not updated by this run (Failed). Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- TE/Settings: **NotUpdated**; UNITS/TE/2. Calculations/Settings Data.xlsx; last saved UTC: 2026-10-25T02:48:05.0869073Z. Not updated by this run (Blocked). Prerequisite TE/AllocationInput did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- TE/DemandInput: **NotUpdated**; UNITS/TE/1. Input/2-DemandExtract.xlsx; last saved UTC: 2026-10-26T01:34:07.4839372Z. Not updated by this run (Blocked). Prerequisite TE/Settings did not complete.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- TE/Workers: **NotUpdated**; UNITS/TE/1. Input/Worker Reconciliation.xlsx; last saved UTC: 2026-10-26T01:31:22.7236830Z. Not updated by this run (Blocked). Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Availability: **NotUpdated**; UNITS/TE/2. Calculations/Capacity-ShiftAvailability.xlsx; last saved UTC: 2026-10-26T01:32:37.8389541Z. Not updated by this run (Blocked). Prerequisite TE/Workers did not complete.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Staff: **NotUpdated**; UNITS/TE/2. Calculations/StaffListMaster.xlsx; last saved UTC: 2026-10-25T02:50:43.8480388Z. Not updated by this run (Blocked). Prerequisite TE/Availability did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Intervals: **NotUpdated**; UNITS/TE/2. Calculations/Intervals.xlsx; last saved UTC: 2026-10-25T02:51:41.5085356Z. Not updated by this run (Blocked). Prerequisite TE/AllocationInput did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/DemandIntervals: **NotUpdated**; UNITS/TE/2. Calculations/DemandIntervals.xlsx; last saved UTC: 2026-10-25T02:53:06.8513510Z. Not updated by this run (Blocked). Prerequisite TE/Intervals did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Demand: **NotUpdated**; UNITS/TE/2. Calculations/Demand.xlsx; last saved UTC: 2026-10-25T02:53:26.6322501Z. Not updated by this run (Blocked). Prerequisite TE/DemandIntervals did not complete.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Shifts: **NotUpdated**; UNITS/TE/2. Calculations/Shifts.xlsx; last saved UTC: 2026-10-25T02:52:48.1964737Z. Not updated by this run (Blocked). Prerequisite TE/Intervals did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/ShiftAverage: **NotUpdated**; UNITS/TE/2. Calculations/AllocationByShiftAverage.xlsx; last saved UTC: 2026-10-25T02:53:31.5351824Z. Not updated by this run (Blocked). Prerequisite TE/Shifts did not complete.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Allocation: **NotUpdated**; UNITS/TE/2. Calculations/Allocation.xlsx; last saved UTC: 2026-10-25T02:53:51.1449193Z. Not updated by this run (Blocked). Prerequisite TE/ShiftAverage did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Distribution: **NotUpdated**; UNITS/TE/2. Calculations/Shift-StaffDistribution.xlsx; last saved UTC: 2026-10-25T02:54:12.0663483Z. Not updated by this run (Blocked). Prerequisite TE/Allocation did not complete.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
- TE/Role:AIN/1: **NotUpdated**; UNITS/TE/2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-25T02:56:05.1871888Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:AIN/2: **NotUpdated**; UNITS/TE/2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-25T02:56:54.2620945Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Demand did not complete.
- TE/Role:AIN/3: **NotUpdated**; UNITS/TE/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-25T02:58:33.3487029Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Demand did not complete.
- TE/Role:RN/1: **NotUpdated**; UNITS/TE/2. Calculations/RN/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-25T02:55:21.8750495Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:RN/2: **NotUpdated**; UNITS/TE/2. Calculations/RN/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-25T02:56:10.4514418Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Demand did not complete.
- TE/Role:RN/3: **NotUpdated**; UNITS/TE/2. Calculations/RN/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-25T02:57:14.5072742Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Demand did not complete.
- TE/Role:EN/1: **NotUpdated**; UNITS/TE/2. Calculations/EN/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-25T02:55:19.4751740Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:EN/2: **NotUpdated**; UNITS/TE/2. Calculations/EN/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-25T02:56:08.1315237Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Demand did not complete.
- TE/Role:EN/3: **NotUpdated**; UNITS/TE/2. Calculations/EN/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-25T02:57:05.1416812Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Demand did not complete.
- TE/Role:CS/1: **NotUpdated**; UNITS/TE/2. Calculations/CS/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-26T00:49:35.0723680Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
- TE/Role:CS/2: **NotUpdated**; UNITS/TE/2. Calculations/CS/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-26T00:50:05.6181868Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Demand did not complete.
- TE/Role:CS/3: **NotUpdated**; UNITS/TE/2. Calculations/CS/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-26T00:50:45.4672090Z. Not updated by this run (Blocked). Prerequisite TE/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Demand did not complete.
- TE/Capacity: **NotUpdated**; UNITS/TE/2. Calculations/Capacity.xlsx; last saved UTC: 2026-10-26T01:05:42.8307654Z. Not updated by this run (Blocked). Prerequisite TE/Role:AIN/3 did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Demand did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Demand did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Demand did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Demand did not complete.
- TE/Effort: **NotUpdated**; UNITS/TE/2. Calculations/Effort.xlsx; last saved UTC: 2026-10-22T01:24:07.4513502Z. Not updated by this run (Blocked). Prerequisite TE/Capacity did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: TE/Settings: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Workers: Prerequisite TE/AllocationInput did not complete.
  Root cause: TE/Shifts: Prerequisite TE/Intervals did not complete.
  Root cause: TE/Role:AIN/1: Prerequisite TE/Demand did not complete.
  Root cause: TE/Role:EN/1: Prerequisite TE/Demand did not complete.
  Root cause: TE/Role:RN/1: Prerequisite TE/Demand did not complete.
  Root cause: TE/Role:CS/1: Prerequisite TE/Demand did not complete.
  Root cause: TE/Allocation: Prerequisite TE/ShiftAverage did not complete.
  Root cause: TE/Distribution: Prerequisite TE/Allocation did not complete.
- JH-RY/AllocationInput: **Verified**; UNITS/JH-RY/1. Input/1-AllocationExtracted.xlsx; last saved UTC: 2026-10-26T01:50:03.7008245Z. Confirmed run output and all recorded declared input metadata still match.
- JH-RY/Settings: **Verified**; UNITS/JH-RY/2. Calculations/Settings Data.xlsx; last saved UTC: 2026-10-26T01:50:23.5390100Z. Confirmed run output and all recorded declared input metadata still match.
- JH-RY/DemandInput: **NotUpdated**; UNITS/JH-RY/1. Input/2-DemandExtract.xlsx; last saved UTC: 2026-10-26T01:34:07.4839372Z. Not updated by this run (Blocked). Prerequisite Date/DemandMaster did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/Workers: **Verified**; UNITS/JH-RY/1. Input/Worker Reconciliation.xlsx; last saved UTC: 2026-10-26T01:51:24.7603175Z. Confirmed run output and all recorded declared input metadata still match.
- JH-RY/Availability: **NotUpdated**; UNITS/JH-RY/2. Calculations/Capacity-ShiftAvailability.xlsx; last saved UTC: 2026-10-26T01:32:37.8389541Z. Not updated by this run (Failed). Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK. The output predates a current declared input.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
- JH-RY/Staff: **NotUpdated**; UNITS/JH-RY/2. Calculations/StaffListMaster.xlsx; last saved UTC: 2026-10-25T02:50:29.4561426Z. Not updated by this run (Blocked). Prerequisite JH-RY/Availability did not complete. The output predates a current declared input.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
- JH-RY/Intervals: **NotUpdated**; UNITS/JH-RY/2. Calculations/Intervals.xlsx; last saved UTC: 2026-10-25T02:51:43.1804577Z. Not updated by this run (Blocked). Prerequisite JH-RY/DemandInput did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/DemandIntervals: **NotUpdated**; UNITS/JH-RY/2. Calculations/DemandIntervals.xlsx; last saved UTC: 2026-10-25T02:53:12.5828427Z. Not updated by this run (Blocked). Prerequisite JH-RY/Intervals did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/Demand: **NotUpdated**; UNITS/JH-RY/2. Calculations/Demand.xlsx; last saved UTC: 2026-10-25T02:53:34.1864864Z. Not updated by this run (Blocked). Prerequisite JH-RY/DemandIntervals did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/Shifts: **NotUpdated**; UNITS/JH-RY/2. Calculations/Shifts.xlsx; last saved UTC: 2026-10-25T02:53:11.1705246Z. Not updated by this run (Blocked). Prerequisite JH-RY/Intervals did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/ShiftAverage: **NotUpdated**; UNITS/JH-RY/2. Calculations/AllocationByShiftAverage.xlsx; last saved UTC: 2026-10-25T02:53:55.1793898Z. Not updated by this run (Blocked). Prerequisite JH-RY/Shifts did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/Allocation: **NotUpdated**; UNITS/JH-RY/2. Calculations/Allocation.xlsx; last saved UTC: 2026-10-25T02:54:22.7184266Z. Not updated by this run (Blocked). Prerequisite JH-RY/ShiftAverage did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/Distribution: **NotUpdated**; UNITS/JH-RY/2. Calculations/Shift-StaffDistribution.xlsx; last saved UTC: 2026-10-25T02:54:50.2759309Z. Not updated by this run (Blocked). Prerequisite JH-RY/Allocation did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
- JH-RY/Role:AIN/1: **NotUpdated**; UNITS/JH-RY/2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-25T02:56:50.3154684Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
- JH-RY/Role:AIN/2: **NotUpdated**; UNITS/JH-RY/2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-25T02:57:38.5377971Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:AIN/3: **NotUpdated**; UNITS/JH-RY/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-25T02:59:14.0025399Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:RN/1: **NotUpdated**; UNITS/JH-RY/2. Calculations/RN/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-25T02:55:53.3085978Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
- JH-RY/Role:RN/2: **NotUpdated**; UNITS/JH-RY/2. Calculations/RN/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-25T02:56:33.2130117Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:RN/3: **NotUpdated**; UNITS/JH-RY/2. Calculations/RN/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-25T02:57:24.4808077Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:EN/1: **NotUpdated**; UNITS/JH-RY/2. Calculations/EN/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-25T02:55:59.8921404Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
- JH-RY/Role:EN/2: **NotUpdated**; UNITS/JH-RY/2. Calculations/EN/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-25T02:56:43.4656926Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:EN/3: **NotUpdated**; UNITS/JH-RY/2. Calculations/EN/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-25T02:57:37.7368139Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:CS/1: **NotUpdated**; UNITS/JH-RY/2. Calculations/CS/CapacityDistrib(A.1)-shifts.xlsx; last saved UTC: 2026-10-26T00:49:35.0723680Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
- JH-RY/Role:CS/2: **NotUpdated**; UNITS/JH-RY/2. Calculations/CS/CapacityDistrib(A.2)-shifts.xlsx; last saved UTC: 2026-10-26T00:50:05.6181868Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Role:CS/3: **NotUpdated**; UNITS/JH-RY/2. Calculations/CS/CapacityDistrib(B)-shifts.xlsx; last saved UTC: 2026-10-26T00:50:45.4672090Z. Not updated by this run (Blocked). Prerequisite JH-RY/Demand did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Capacity: **NotUpdated**; UNITS/JH-RY/2. Calculations/Capacity.xlsx; last saved UTC: 2026-10-26T01:05:42.8307654Z. Not updated by this run (Blocked). Prerequisite JH-RY/Role:AIN/3 did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Demand did not complete.
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Demand did not complete.
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Demand did not complete.
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Demand did not complete.
- JH-RY/Effort: **NotUpdated**; UNITS/JH-RY/2. Calculations/Effort.xlsx; last saved UTC: 2026-10-25T04:02:40.1784277Z. Not updated by this run (Blocked). Prerequisite JH-RY/Capacity did not complete. The output predates a current declared input.
  Root cause: TE/AllocationInput: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'AllocationExtracted (2)/AllocationExtracted' failed foreground refresh: [Expression.Error] Master CS locations require unique Facilities mappings to the complete current Unit.
  Root cause: JH-RY/Availability: Worker exit 1; phase Refreshing; save confirmed: False; Required worksheet query 'ResDayShift/ResDayShift' failed foreground refresh: [ResDayShift validation] Required availability checks failed: Unique worker/date leave totals; Eligible worker roles have roster shifts; Settings shift endpoints; Unique worker shift rows; Segments obey availability and exclusions; Disjoint segments and reconciled hours; Published facility and roles. Inspect ResDayShift_CHECK.
  Root cause: JH-RY/Shifts: Prerequisite JH-RY/Intervals did not complete.
  Root cause: JH-RY/Role:AIN/1: Prerequisite JH-RY/Demand did not complete.
  Root cause: JH-RY/Role:EN/1: Prerequisite JH-RY/Demand did not complete.
  Root cause: JH-RY/Role:RN/1: Prerequisite JH-RY/Demand did not complete.
  Root cause: JH-RY/Role:CS/1: Prerequisite JH-RY/Demand did not complete.
  Root cause: JH-RY/Allocation: Prerequisite JH-RY/ShiftAverage did not complete.
  Root cause: JH-RY/Distribution: Prerequisite JH-RY/Allocation did not complete.
