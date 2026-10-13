# DATExx-Whiddon source file inventory

Updated 13 October 2026, Australia/Sydney.

DATExx-Whiddon currently contains **26 workbook files in its organisation and unit `1. Input` folders**. These include source data, manually maintained inputs and transformed input workbooks. External dependencies include LeaveBalance files, the public path-mapping workbook and Whiddon SharePoint lists.

All paths are relative to the ResidentialCare repository unless a table states that they are relative to the DATExx-Whiddon client folder. Timestamped backups, extracted M files and packaged report data are excluded from the input-workbook total. Import references describe the available extracted M text; they do not establish the current definitions inside Excel workbooks.

## Remaining input files

| Scope | Folder under `CLIENT/DATExx-Whiddon/` | Workbook count |
| --- | --- | --- |
| Organisation | `1. Input/` | 3 |
| BD | `UNITS/BD/1. Input/` | 8 |
| JH-RY | `UNITS/JH-RY/1. Input/` | 8 |
| TE | `UNITS/TE/1. Input/` | 7 |

### Organisation

Files below are relative to `CLIENT/DATExx-Whiddon/1. Input/`.

| File | Extracted M source beside workbook |
| --- | --- |
| `1-AllocationExtracted.xlsx` | Present |
| `Demand-MasterRoster Manual Read.xlsx` | Present |
| `Worker Reconciliation.xlsx` | Present |

### BD

Files below are relative to `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/`.

| File | Extracted M source beside workbook |
| --- | --- |
| `1-AllocationExtracted.xlsx` | Present |
| `2-DemandExtract.xlsx` | Present |
| `Allocation/Published Roster.xlsx` | Absent; contents or imports not inferred |
| `Occupancy care minute -roster meeting.xlsx` | Absent; contents or imports not inferred |
| `StaffList Availability.xlsx` | Absent; contents or imports not inferred |
| `Worker Reconciliation.xlsx` | Present |
| `Worker's Report.xlsx` | Absent; contents or imports not inferred |
| `whiddon_availability_leave_extraction.xlsx` | Absent; contents or imports not inferred |

### JH-RY

Files below are relative to `CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/`.

| File | Extracted M source beside workbook |
| --- | --- |
| `1-AllocationExtracted.xlsx` | Present |
| `2-DemandExtract.xlsx` | Present |
| `Allocation/Published Roster.xlsx` | Absent; contents or imports not inferred |
| `Occupancy care minute -roster meeting.xlsx` | Absent; contents or imports not inferred |
| `StaffList Availability.xlsx` | Absent; contents or imports not inferred |
| `Worker Reconciliation.xlsx` | Present |
| `Worker's Report.xlsx` | Absent; contents or imports not inferred |
| `whiddon_availability_leave_extraction.xlsx` | Absent; contents or imports not inferred |

### TE

Files below are relative to `CLIENT/DATExx-Whiddon/UNITS/TE/1. Input/`.

| File | Extracted M source beside workbook |
| --- | --- |
| `1-AllocationExtracted.xlsx` | Present |
| `2-DemandExtract.xlsx` | Present |
| `Allocation/Published Roster.xlsx` | Absent; contents or imports not inferred |
| `StaffList Availability.xlsx` | Absent; contents or imports not inferred |
| `Worker Reconciliation.xlsx` | Present |
| `Worker's Report.xlsx` | Absent; contents or imports not inferred |
| `whiddon_availability_leave_extraction.xlsx` | Absent; contents or imports not inferred |

## Recorded raw source imports

References below come from retained workbooks with exact adjacent extracted M files. `Unit1Path` and similarly named path queries are recorded as written; dynamic roots have not been evaluated.

| Source or path expression | Importing workbook and query |
| --- | --- |
| `Unit1Path & "\1. Input\Allocation\Published Roster.xlsx"` | `1. Input/1-AllocationExtracted.xlsx_PowerQuery.m:276` (IMPORT Roster)<br>`1. Input/Worker Reconciliation.xlsx_PowerQuery.m:114` (IMPORT Combined Roster)<br>`UNITS/BD/1. Input/1-AllocationExtracted.xlsx_PowerQuery.m:276` (IMPORT Roster)<br>`UNITS/BD/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:114` (IMPORT Combined Roster)<br>`UNITS/JH-RY/1. Input/1-AllocationExtracted.xlsx_PowerQuery.m:276` (IMPORT Roster)<br>`UNITS/JH-RY/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:114` (IMPORT Combined Roster)<br>`UNITS/TE/1. Input/1-AllocationExtracted.xlsx_PowerQuery.m:276` (IMPORT Roster)<br>`UNITS/TE/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:114` (IMPORT Combined Roster) |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Master Roster.xlsx` | `1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m:15` (IMPORT Master) |
| `Unit1Path & "\1. Input\whiddon_availability_leave_extraction.xlsx"` | `1. Input/Worker Reconciliation.xlsx_PowerQuery.m:127` (IMPORT Combined Availabilities)<br>`UNITS/BD/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:127` (IMPORT Combined Availabilities)<br>`UNITS/JH-RY/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:127` (IMPORT Combined Availabilities)<br>`UNITS/TE/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:127` (IMPORT Combined Availabilities) |
| `Unit1Path & "\1. Input\Worker's Report.xlsx"` | `1. Input/Worker Reconciliation.xlsx_PowerQuery.m:413` (IMPORT Current_Workers)<br>`UNITS/BD/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:413` (IMPORT Current_Workers)<br>`UNITS/JH-RY/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:413` (IMPORT Current_Workers)<br>`UNITS/TE/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:413` (IMPORT Current_Workers) |
| `#"FilePath - 1Input" & "\Worker's Report.xlsx"` | `UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m:151` (IMPORT Worker's Report)<br>`UNITS/JH-RY/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m:151` (IMPORT Worker's Report)<br>`UNITS/TE/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m:134` (IMPORT Worker's Report) |
| `#"FilePath - 1Input" & "\whiddon_availability_leave_extraction.xlsx"` | `UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m:353` (IMPORT Availability Leave Source)<br>`UNITS/JH-RY/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m:353` (IMPORT Availability Leave Source)<br>`UNITS/TE/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m:336` (IMPORT Availability Leave Source) |
| `#"FilePath-1Input" & "\Worker's Report.xlsx"` | `UNITS/BD/2. Calculations/StaffListMaster.xlsx_PowerQuery.m:312` (IMPORT StaffList Worker's Report)<br>`UNITS/TE/2. Calculations/StaffListMaster.xlsx_PowerQuery.m:313` (IMPORT StaffList Worker's Report) |

## Controls outside input folders

| Repository relative file | Recorded use |
| --- | --- |
| `CLIENT/DATExx-Whiddon/2. Calculations/EOW/Boundary Intersections.xlsx` | Present; direct Power Query import not found |
| `CLIENT/DATExx-Whiddon/2. Calculations/EOW/Grid Thresholds.xlsx` | Imported by EOW/2DRead.xlsx |
| `CLIENT/DATExx-Whiddon/Roles.xlsx` | Present; direct Power Query import not found |
| `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Settings Data.xlsx` | Unit calculation and demand controls |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Settings Data.xlsx` | Unit calculation and demand controls |
| `CLIENT/DATExx-Whiddon/UNITS/TE/2. Calculations/Settings Data.xlsx` | Unit calculation and demand controls |

## External files referenced by Power Query

| Repository relative file | File status | Reference evidence |
| --- | --- | --- |
| `../../../Public/Public Scripts/CentriSyncPaths.xlsx` | Present | `1. Input/1-AllocationExtracted.xlsx_PowerQuery.m:80` (PathTABLE)<br>`1. Input/Worker Reconciliation.xlsx_PowerQuery.m:11` (IMPORT CentriSyncPaths)<br>`2. Calculations/Cost/Cost..xlsx_PowerQuery.m:11` (IMPORT CentriSyncPaths)<br>`2. Calculations/E-O-I/Effort-All.xlsx_PowerQuery.m:10` (IMPORT CentriSyncPaths)<br>and 78 further references |
| `../LeaveBalance/Workflows/LeaveBalance/Analysis/1.Input/Leave Balances Report.xlsx` | Present | `1. Input/Worker Reconciliation.xlsx_PowerQuery.m:139` (IMPORT Leave Balances)<br>`UNITS/BD/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:139` (IMPORT Leave Balances)<br>`UNITS/JH-RY/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:139` (IMPORT Leave Balances)<br>`UNITS/TE/1. Input/Worker Reconciliation.xlsx_PowerQuery.m:139` (IMPORT Leave Balances) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/Allocation.xlsx` | Not found at referenced path | `2. Calculations/Change/AllocationChange.xlsx_PowerQuery.m:58` (Allocation) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/SETTINGS.xlsx` | Not found at referenced path | `2. Calculations/Change/AllocationChange.xlsx_PowerQuery.m:72` (ShiftDuration)<br>`2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:53` (ShiftDuration)<br>`2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:172` (PermutationDimensions)<br>`2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:55` (ShiftDuration)<br>and 1 further references |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/2. Calculations/E-O-I/Effort.xlsx` | Not found at referenced path | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:31` (EffortAllMatrixAG1_1D1)<br>`2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:33` (EffortAllMatrixAG1_1D1) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/2. Calculations/SETTINGS.xlsx` | Not found at referenced path | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:38` (Shifts)<br>`2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:40` (Shifts) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/ResidualAllocation-Facility.xlsx` | Not found at referenced path | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:74` (Table_Master_Avail_Alloc_ResidualMAIT) |

## LeaveBalance sources and reporting dependencies

`pq.project.json` configures `../LeaveBalance/Workflows/LeaveBalance/Analysis/2.Calulations` as the LeaveBalanceStats workbook root. The batch catalogue names `LBStats.xlsx` and `TableauConnect-LB.xlsx`. The LeaveBalance lineage document records `Leave Balances Report.xlsx` and `Worker's Report.xlsx` as LBStats inputs. Other files listed here are present in that input folder; their use by ResidentialCare is not established.

| File under `../LeaveBalance/Workflows/LeaveBalance/Analysis/1.Input/` | Status |
| --- | --- |
| `Leave Balances Report.xlsx` | Present |
| `Leave Booked Report - 12 months forward.xlsx` | Present |
| `Leave Taken Report - past 12 months.xlsx` | Present |
| `Published Roster.xlsx` | Present |
| `timesheets.xlsx` | Present |
| `Whiddon Worker Reconciliation.xlsx` | Present |
| `whiddon_availability_leave_extraction.xlsx` | Present |
| `Worker's Report.xlsx` | Present |

| Configured external calculation workbook | Status |
| --- | --- |
| `../LeaveBalance/Workflows/LeaveBalance/Analysis/2.Calulations/LBStats.xlsx` | Present |
| `../LeaveBalance/Workflows/LeaveBalance/Analysis/2.Calulations/TableauConnect-LB.xlsx` | Present |

## SharePoint sources

Extracted M refers to `https://centri001.sharepoint.com/sites/WhiddonCENTRI` through `SharePoint.Tables`. The four recorded list IDs serve Roles (`6b0b0767-44f4-4bb8-b116-f1e99b3476f0`), RoleAnalysis (`bad3beb8-7064-4e11-9edd-d62dac5d702d`), Facilities (`1987eadb-ad2e-491a-a927-e5585667d4c5`) and Facility or Unit Analysis (`17ed8e30-5707-4af2-8e21-d68eb8184287`). Cloud availability has not been checked.

## Imported file families

This includes intermediate calculation workbooks used by downstream workbooks. Counts are occurrences in extracted source text, including repeated imports; they are not refresh counts.

| Referenced filename | Import occurrences | Importing extracted source files |
| --- | --- | --- |
| `1-AllocationExtracted.xlsx` | 18 | 18 |
| `2-DemandExtract.xlsx` | 5 | 5 |
| `Allocation.xlsx` | 3 | 3 |
| `AllocationByShiftAverage.xlsx` | 16 | 14 |
| `AllocationChange.xlsx` | 1 | 1 |
| `Budget and AN-ACC Comparison.xlsx` | 2 | 2 |
| `Capacity-ShiftAvailability.xlsx` | 15 | 15 |
| `Capacity.xlsx` | 2 | 2 |
| `CapacityDistrib(A.1)-shifts.xlsx` | 81 | 27 |
| `CapacityDistrib(A.2)-shifts.xlsx` | 21 | 15 |
| `CapacityDistrib(B)-shifts.xlsx` | 18 | 3 |
| `CentriSyncPaths.xlsx` | 82 | 82 |
| `Changes.xlsm` | 1 | 1 |
| `Demand-MasterRoster Manual Read.xlsx` | 4 | 4 |
| `Demand.xlsx` | 14 | 14 |
| `DemandIntervals.xlsx` | 2 | 2 |
| `Effort-All.xlsx` | 2 | 2 |
| `Effort.xlsx` | 5 | 3 |
| `EffortOutcomeLogXY.xlsx` | 2 | 1 |
| `EffortOutcomes.xlsx` | 4 | 2 |
| `Grid Thresholds.xlsx` | 1 | 1 |
| `Inefficiencies.xlsx` | 2 | 2 |
| `Intervals.xlsx` | 11 | 8 |
| `Leave Balances Report.xlsx` | 4 | 4 |
| `Master Roster.xlsx` | 1 | 1 |
| `Published Roster.xlsx` | 8 | 8 |
| `ResidualAllocation-Facility.xlsx` | 5 | 2 |
| `Settings Data.xlsx` | 73 | 53 |
| `SETTINGS.xlsx` | 8 | 3 |
| `Shifts.xlsx` | 5 | 5 |
| `StaffListMaster.xlsx` | 29 | 28 |
| `StafMasterList-All.xlsx` | 1 | 1 |
| `whiddon_availability_leave_extraction.xlsx` | 7 | 7 |
| `Worker Reconciliation.xlsx` | 3 | 3 |
| `Worker's Report.xlsx` | 9 | 9 |

## Legacy and unresolved references

The following source references remain recorded. The presence of a query does not establish that it is loaded or evaluated. `INPUT RosterBudget-OFF` queries reference `Budget and AN-ACC Comparison.xlsx`, which is absent from the client file inventory. Legacy `Folder` and `Folder1` roots cannot be resolved from source text alone.

| Referenced file or expression | Reference evidence |
| --- | --- |
| `Folder & "\MAIT\2. Calculations\Changes.xlsm"` | `2. Calculations/Change/AllocationChange.xlsx_PowerQuery.m:39` (Table_ChangesAllRoles) |
| `Folder & "\2. Calculations\SETTINGS.xlsx"` | `2. Calculations/Change/AllocationChange.xlsx_PowerQuery.m:50` (Shifts) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/SETTINGS.xlsx` | `2. Calculations/Change/AllocationChange.xlsx_PowerQuery.m:72` (ShiftDuration) |
| `Folder &  "\MAIT\2. Calculations\ResidualAllocation-Facility.xlsx"` | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:12` (Table_MinResidualHRSResDayShiftMAITRaw) |
| `Folder &  "\MAIT\2. Calculations\ResidualAllocation-Facility.xlsx"` | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:22` (Table_MinResidualHRSResDayShiftMAIT) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/2. Calculations/SETTINGS.xlsx` | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:38` (Shifts) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/SETTINGS.xlsx` | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:53` (ShiftDuration) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/ResidualAllocation-Facility.xlsx` | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:74` (Table_Master_Avail_Alloc_ResidualMAIT) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/SETTINGS.xlsx` | `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx_PowerQuery.m:172` (PermutationDimensions) |
| `Folder &  "\MAIT\2. Calculations\ResidualAllocation-Facility.xlsx"` | `2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:12` (Table_MinResidualHRSResDayShiftMAITRaw) |
| `Folder &  "\MAIT\2. Calculations\ResidualAllocation-Facility.xlsx"` | `2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:23` (Table_MinResidualHRSResDayShiftMAIT) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/2. Calculations/SETTINGS.xlsx` | `2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:40` (Shifts) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/SETTINGS.xlsx` | `2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:55` (ShiftDuration) |
| `../../Centri/2. Accounts - 2.0 RFBI/04.1 PoC/04 Analysis/230714/MAIT/2. Calculations/SETTINGS.xlsx` | `2. Calculations/Change/SS/ResidualAllocation.xlsx_PowerQuery.m:172` (PermutationDimensions) |
| `Folder &  "\1. Input\Budget and AN-ACC Comparison.xlsx"` | `UNITS/BD/2. Calculations/Demand.xlsx_PowerQuery.m:130` (INPUT RosterBudget-OFF) |
| `Folder &  "\1. Input\Budget and AN-ACC Comparison.xlsx"` | `UNITS/TE/2. Calculations/Demand.xlsx_PowerQuery.m:130` (INPUT RosterBudget-OFF) |

## Tableau file connections

These are saved path references in unpackaged `.twb` files. They include other machines, other clients and report extracts. A similar local filename has not been substituted for a missing recorded path. Packaged `.twbx` contents are outside this inventory.

| Repository relative referenced file | File status | Referencing report |
| --- | --- | --- |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/2. Calculations/Cost/Cost..xlsx` | Not found at recorded path | `3. Report/Cost-LT_v2025.3.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/2. Calculations/E-O-I/Effort-All.xlsx` | Not found at recorded path | `3. Report/Availability Redistribution.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/2. Calculations/EOW/Grid Thresholds.xlsx` | Not found at recorded path | `3. Report/EOW v07 - AG1-1DS.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/3. Report/ShiftProfilex1F-LT.twb Files/Data/2 1. calculations/AllocationByShiftAverage.xlsx` | Not found at recorded path | `3. Report/ShiftProfilex1F-LT.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/3. Report/ShiftProfilex1F-LT.twb Files/Data/2 1. calculations/Effort.xlsx` | Not found at recorded path | `3. Report/ShiftProfilex1F-LT.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/3. Report/ShiftProfilex1F-LT.twb Files/Data/2. Calculations/Effort.xlsx` | Not found at recorded path | `3. Report/ShiftProfilex1F-LT.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/3. Report/ShiftProfilex1F-LT.twb Files/Data/Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/UNITS/Unit1/2. Calculations/Intervals.xlsx` | Not found at recorded path | `3. Report/ShiftProfilex1F-LT.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/3. Report/ShiftProfilex1F-LT.twb Files/Data/Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/UNITS/Unit1/2. Calculations/Shift-StaffDistribution.xlsx` | Not found at recorded path | `3. Report/ShiftProfilex1F-LT.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/3. Report/ShiftProfilex1F-LT.twb Files/Data/Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/UNITS/Unit1/2. Calculations/Shifts.xlsx` | Not found at recorded path | `3. Report/ShiftProfilex1F-LT.twb` |
| `../../../Cliff's Computer/Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/ResidentialCare/CLIENT2/DATExx/UNITS/Unit1/2. Calculations/StaffListMaster.xlsx` | Not found at recorded path | `3. Report/Availability Redistribution.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/2. Calculations/Change/Allocation.xlsx` | Not found at recorded path | `3. Report/EOW v07 - AG1-1DS.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/2. Calculations/Cost/Cost..xlsx` | Not found at recorded path | `3. Report/Cost-LT_v2025.3.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/2. Calculations/E-O-I/Effort-All.xlsx` | Not found at recorded path | `3. Report/Availability Redistribution.twb`<br>`3. Report/Availability RedistributionY.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/2. Calculations/E-O-I/StafMasterList-All.xlsx` | Not found at recorded path | `3. Report/Availability RedistributionY.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/2. Calculations/EOW/2DRead.xlsx` | Not found at recorded path | `3. Report/EOW v07 - AG1-1DS.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/2. Calculations/EOW/Grid Thresholds.xlsx` | Not found at recorded path | `3. Report/EOW v07 - AG1-1DS.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/3. Report/Availability Effort (Effort-All).hyper` | Not found at recorded path | `3. Report/Availability RedistributionY.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/3. Report/CostAG1_1DayShiftAB (Cost).hyper` | Not found at recorded path | `3. Report/Cost-LT_v2025.3.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/3. Report/OutcomeXYall- OutAG1-1DS (2DRead)y.hyper` | Not found at recorded path | `3. Report/EOW v07 - AG1-1DS.twb` |
| `../../../Cliff's Computer/CENTRINotSync/mcode/ResidentialCare/CLIENT/DATExx/3. Report/Thresholds (Grid Thresholds).hyper` | Not found at recorded path | `3. Report/EOW v07 - AG1-1DS.twb` |
| `../../Centri/3.0 Product - Documents/AEM/Capacity Distribution/AC Data 2406/2. Calculations/E-O-I/Effort-All.xlsx` | Not found at recorded path | `3. Report/Availability Redistribution.twb`<br>`3. Report/Availability RedistributionY.twb` |
| `../../Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/01 Christadelphian Aged Care/240715/FACILITIES/ASHB/2. Calculations/Effort.xlsx` | Present | `3. Report/ShiftProfile-BD.twb`<br>`3. Report/ShiftProfile-JH-RY.twb` |
| `../../Centri/4. Production - Documents/WFEffectiveness/4.1.1 AGED CARE/HomeCare/KA Aged Care/2. Route Analysis/260701/2. Calculations/1Pool (Tableau Connection).hyper` | Not found at recorded path | `3. Report/Effort and Inefficiencies-temp.twb`<br>`3. Report/Effort and Inefficiencies-X.twb`<br>`3. Report/Effort and Inefficiencies.twb` |
| `../../Resrodel/Customers - Documents/IndoChinese Aged Care/4. Analysis/221219/5.Tableau Data Extract/EffortOutcomes3D2D (2DRead).hyper` | Not found at recorded path | `3. Report/EOW v07 - AG1-1DS.twb` |
| `../LeaveBalance/Workflows/LeaveBalance/Analysis/2.Calulations/TableauConnect-LB.xlsx` | Present | `3. Report/ShiftProfile-JH-RY.twb` |
| `CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All - X.xlsx` | Present | `3. Report/Effort and Inefficiencies-X.twb` |
| `CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All.xlsx` | Present | `3. Report/Effort and Inefficiencies-temp.twb`<br>`3. Report/Effort and Inefficiencies.twb` |
| `CLIENT/DATExx-Whiddon/2. Calculations/Tableau Connection - X.xlsx` | Present | `3. Report/Effort and Inefficiencies-X.twb` |
| `CLIENT/DATExx-Whiddon/2. Calculations/Tableau Connection.xlsx` | Present | `3. Report/Effort and Inefficiencies-temp.twb`<br>`3. Report/Effort and Inefficiencies-X.twb`<br>`3. Report/Effort and Inefficiencies.twb`<br>`3. Report/ShiftProfile-JH-RY.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AllocationByShiftAverage.xlsx` | Present | `3. Report/ShiftProfile-BD.twb`<br>`3. Report/ShiftProfile-JH-RY.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Effort.xlsx` | Present | `3. Report/ShiftProfile-BD.twb`<br>`3. Report/ShiftProfile-JH-RY.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shift-StaffDistribution.xlsx` | Present | `3. Report/ShiftProfile-BD.twb`<br>`3. Report/ShiftProfile-JH-RY.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shifts.xlsx` | Present | `3. Report/ShiftProfile-BD.twb`<br>`3. Report/ShiftProfile-JH-RY.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Shift-StaffDistribution.xlsx` | Present | `3. Report/ShiftProfile-TE.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Shifts.xlsx` | Present | `3. Report/ShiftProfile-TE.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/TE/2. Calculations/AllocationByShiftAverage.xlsx` | Present | `3. Report/ShiftProfile-TE.twb` |
| `CLIENT/DATExx-Whiddon/UNITS/TE/2. Calculations/Effort.xlsx` | Present | `3. Report/ShiftProfile-TE.twb` |

## Retained imports into deleted unit folders

3 recorded file imports in retained workbook sources still point into the deleted unit folders. These references were not changed by the requested deletions.

| Deleted target | Retained importer | Query | File status |
| --- | --- | --- | --- |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Master Roster.xlsx` | `CLIENT/DATExx-Whiddon/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m:15` | IMPORT Master | Missing after unit-folder deletion |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/Settings Data.xlsx` | `CLIENT/DATExx-Whiddon/2. Calculations/Cost/Cost..xlsx_PowerQuery.m:118` | IMPORT Settings Data | Missing after unit-folder deletion |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/StaffListMaster.xlsx` | `CLIENT/DATExx-Whiddon/2. Calculations/Tableau Connection.xlsx_PowerQuery.m:144` | IMPORT Table_Masterlist | Missing after unit-folder deletion |

## Source coverage

89 retained workbook and extracted M pairs were available. Files without extracted M have only been inventoried by filename; their contents and imports are not inferred.

## Requested deletions

28 individually named files were deleted. `CLIENT/DATExx-Whiddon/UNITS/Unit1` and `CLIENT/DATExx-Whiddon/UNITS/Unit2` were also deleted, containing 311 additional files at deletion time.

| Deleted repository relative file |
| --- |
| `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/2-DemandExtract copy.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Allocation/Ashburn Working roster 7 to 20 October 24.xls` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Allocation/Master Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Allocation/Worker's Report.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/AN-ACC Model ASH 2410.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Demand-MasterRoster Manual ReadBD.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Published Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/2-DemandExtract copy.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/Allocation/Ashburn Working roster 7 to 20 October 24.xls` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/Allocation/Master Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/Allocation/Worker's Report.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/AN-ACC Model ASH 2410.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/Published Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/TE/1. Input/2-DemandExtract copy.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/TE/1. Input/Allocation/Ashburn Working roster 7 to 20 October 24.xls` |
| `CLIENT/DATExx-Whiddon/UNITS/TE/1. Input/Allocation/Master Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/TE/1. Input/Allocation/Worker's Report.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/TE/1. Input/Published Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Allocation/Ashburn Working roster 7 to 20 October 24.xls` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Allocation/Master Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Allocation/Worker's Report.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/AN-ACC Model ASH 2410.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Published Roster.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit2/1. Input/Allocation/Ashburn Working roster 7 to 20 October 24.xls` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit2/1. Input/Allocation/Worker's Report.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit2/1. Input/AN-ACC Model ASH 2410.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit2/1. Input/Demand-MasterRoster Manual ReadX.xlsx` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit2/1. Input/Published Roster.xlsx` |

| Deleted unit folder |
| --- |
| `CLIENT/DATExx-Whiddon/UNITS/Unit1` |
| `CLIENT/DATExx-Whiddon/UNITS/Unit2` |

Excel workbook contents were not inspected, and no synchronization or refresh was performed. The source report records the resulting filesystem state and available text references.
