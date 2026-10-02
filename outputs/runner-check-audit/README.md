# ResidentialCare runner check audit

This is a source-only audit of the non-role workbooks in the current batch catalogue plus one representative three-workbook role folder. It does not inspect or alter an Excel workbook.

## Result

- Catalogue workbooks: 138
- Workbooks in this audit: 96
- Representative role folder: TE/AIN (three files)
- Equivalent role-folder workbook instances omitted: 42
- Exact extracted M sources audited: 62
- Workbooks blocked from source audit by missing exact extracted M: 34
- Default-run workbooks in this audit: 62; source-backed: 30; missing source: 32
- Source-backed workbooks with named checks/validation/diagnostics: 33
- Named check-like queries: 153
- Source-backed workbooks with explicit M error gates: 58
- Queries containing explicit M error gates: 240
- Named check-like queries with no source-query consumer: 23

## Runner finding

The expanded runner currently treats refresh-idle plus Excel calculation-done as success. It does not enumerate Power Query checks, interpret returned check tables, scan loaded outputs for Excel errors, or prove that connection-only check queries were demanded. It can therefore save and report success without naming a failed, skipped, or unobserved workbook check.

A safe fail-closed contract needs every required check to have a machine-readable result and an inspectable load surface. A query name alone is not enough: many check queries return tables, some raise M errors, some capture upstream errors, and some are not referenced by another source query.

## Observed runner evidence

- Latest run inspected: `CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh/20260928-065316-5a404270`
- Workbook refresh logs: 79
- Explicit workbook-query/check-result flag lines: 0
- Logs reporting zero worksheet query tables: 67

## Required fail-closed runner contract

1. After refresh and calculation settle, but before Save, evaluate every required workbook check through a declared, machine-readable result surface.
2. Treat a failed check, an Excel/PQ error, a required check that did not run, or a required check that cannot be inspected as a workbook failure. Log the workbook, query/check name, status, count/value, and message.
3. Close the failed workbook without saving, stop new batch dispatch, and mark every not-started job blocked by the check failure. Do not continue unrelated branches after a calculation-integrity failure.
4. Keep advisory diagnostics separate from required gates. Names such as `DIAGNOSTICS` and `TEST` cannot automatically be treated as failures without a declared policy.
5. Extract the missing exact M sources before changing checks in those workbooks. Do not infer their contract from a copy, backup, donor, or same-stem file.

## Missing exact sources

| Default | Job | Workbook | Expected source |
|---|---|---|---|
| No | Unit1/Workers | UNITS/Unit1/1. Input/Worker Reconciliation.xlsx | CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Worker Reconciliation.xlsx_PowerQuery.m |
| No | Unit2/Workers | UNITS/Unit2/1. Input/Worker Reconciliation.xlsx | CLIENT/DATExx-Whiddon/UNITS/Unit2/1. Input/Worker Reconciliation.xlsx_PowerQuery.m |
| Yes | BD/AllocationInput | UNITS/BD/1. Input/1-AllocationExtracted.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/1-AllocationExtracted.xlsx_PowerQuery.m |
| Yes | BD/Settings | UNITS/BD/2. Calculations/Settings Data.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Settings Data.xlsx_PowerQuery.m |
| Yes | BD/DemandMaster | UNITS/BD/1. Input/Demand-MasterRoster Manual Read.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m |
| Yes | BD/DemandInput | UNITS/BD/1. Input/2-DemandExtract.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/2-DemandExtract.xlsx_PowerQuery.m |
| Yes | BD/Workers | UNITS/BD/1. Input/Worker Reconciliation.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Worker Reconciliation.xlsx_PowerQuery.m |
| Yes | BD/Availability | UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m |
| Yes | BD/Staff | UNITS/BD/2. Calculations/StaffListMaster.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/StaffListMaster.xlsx_PowerQuery.m |
| Yes | BD/Intervals | UNITS/BD/2. Calculations/Intervals.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Intervals.xlsx_PowerQuery.m |
| Yes | BD/DemandIntervals | UNITS/BD/2. Calculations/DemandIntervals.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/DemandIntervals.xlsx_PowerQuery.m |
| Yes | BD/Demand | UNITS/BD/2. Calculations/Demand.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Demand.xlsx_PowerQuery.m |
| Yes | BD/Shifts | UNITS/BD/2. Calculations/Shifts.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shifts.xlsx_PowerQuery.m |
| Yes | BD/ShiftAverage | UNITS/BD/2. Calculations/AllocationByShiftAverage.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AllocationByShiftAverage.xlsx_PowerQuery.m |
| Yes | BD/Allocation | UNITS/BD/2. Calculations/Allocation.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Allocation.xlsx_PowerQuery.m |
| Yes | BD/Distribution | UNITS/BD/2. Calculations/Shift-StaffDistribution.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shift-StaffDistribution.xlsx_PowerQuery.m |
| Yes | BD/MultiRole | UNITS/BD/2. Calculations/MutliRoleCheck.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/MutliRoleCheck.xlsx_PowerQuery.m |
| Yes | BD/Effort | UNITS/BD/2. Calculations/Effort.xlsx | CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Effort.xlsx_PowerQuery.m |
| Yes | JH-RY/AllocationInput | UNITS/JH-RY/1. Input/1-AllocationExtracted.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/1-AllocationExtracted.xlsx_PowerQuery.m |
| Yes | JH-RY/Settings | UNITS/JH-RY/2. Calculations/Settings Data.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Settings Data.xlsx_PowerQuery.m |
| Yes | JH-RY/DemandMaster | UNITS/JH-RY/1. Input/Demand-MasterRoster Manual Read.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m |
| Yes | JH-RY/DemandInput | UNITS/JH-RY/1. Input/2-DemandExtract.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/2-DemandExtract.xlsx_PowerQuery.m |
| Yes | JH-RY/Workers | UNITS/JH-RY/1. Input/Worker Reconciliation.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/1. Input/Worker Reconciliation.xlsx_PowerQuery.m |
| Yes | JH-RY/Staff | UNITS/JH-RY/2. Calculations/StaffListMaster.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/StaffListMaster.xlsx_PowerQuery.m |
| Yes | JH-RY/Intervals | UNITS/JH-RY/2. Calculations/Intervals.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Intervals.xlsx_PowerQuery.m |
| Yes | JH-RY/DemandIntervals | UNITS/JH-RY/2. Calculations/DemandIntervals.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/DemandIntervals.xlsx_PowerQuery.m |
| Yes | JH-RY/Demand | UNITS/JH-RY/2. Calculations/Demand.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Demand.xlsx_PowerQuery.m |
| Yes | JH-RY/Shifts | UNITS/JH-RY/2. Calculations/Shifts.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Shifts.xlsx_PowerQuery.m |
| Yes | JH-RY/ShiftAverage | UNITS/JH-RY/2. Calculations/AllocationByShiftAverage.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/AllocationByShiftAverage.xlsx_PowerQuery.m |
| Yes | JH-RY/Allocation | UNITS/JH-RY/2. Calculations/Allocation.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Allocation.xlsx_PowerQuery.m |
| Yes | JH-RY/Distribution | UNITS/JH-RY/2. Calculations/Shift-StaffDistribution.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Shift-StaffDistribution.xlsx_PowerQuery.m |
| Yes | JH-RY/MultiRole | UNITS/JH-RY/2. Calculations/MutliRoleCheck.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/MutliRoleCheck.xlsx_PowerQuery.m |
| Yes | JH-RY/Capacity | UNITS/JH-RY/2. Calculations/Capacity.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Capacity.xlsx_PowerQuery.m |
| Yes | JH-RY/Effort | UNITS/JH-RY/2. Calculations/Effort.xlsx | CLIENT/DATExx-Whiddon/UNITS/JH-RY/2. Calculations/Effort.xlsx_PowerQuery.m |

## Named checks with no source-query consumer

These checks are not demanded by another query in the extracted source. They may be loaded independently, but the current runner neither proves that nor interprets their result.

| Default | Workbook | Query | Raises M error |
|---|---|---|---|
| Yes | 2. Calculations/EOW/2DRead.xlsx | Effort+OutcomesXYImport Check | No |
| Yes | 2. Calculations/EOW/EffortOutcomeLogXY.xlsx | Null Matrix Check | No |
| Yes | UNITS/TE/1. Input/1-AllocationExtracted.xlsx | AllocationExtractedCheck | No |
| Yes | UNITS/TE/1. Input/Demand-MasterRoster Manual Read.xlsx | LocRoleDayShift%CHECK | No |
| Yes | UNITS/TE/2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx | CheckCap | No |
| Yes | UNITS/TE/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx | CheckReAllocation | No |
| Yes | UNITS/TE/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx | ReduceCheck | No |
| Yes | UNITS/TE/2. Calculations/Demand.xlsx | ShiftDemandHCAverageCheck | No |
| Yes | UNITS/TE/2. Calculations/Demand.xlsx | Table_ShiftDemandUnitINTERVALCheck | No |
| Yes | UNITS/TE/2. Calculations/DemandIntervals.xlsx | ShiftDemandUnitINTERVALCheck | No |
| Yes | UNITS/TE/2. Calculations/DemandIntervals.xlsx | Table_ShiftDemandHRSCheck | No |
| No | UNITS/Unit1/1. Input/1-AllocationExtracted.xlsx | AllocationExtractedCheck | No |
| No | UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx | LocRoleDayShift%CHECK | No |
| No | UNITS/Unit1/2. Calculations/Demand.xlsx | ShiftDemandHCAverageCheck | No |
| No | UNITS/Unit1/2. Calculations/Demand.xlsx | Table_ShiftDemandUnitINTERVALCheck | No |
| No | UNITS/Unit1/2. Calculations/DemandIntervals.xlsx | ShiftDemandUnitINTERVALCheck | No |
| No | UNITS/Unit1/2. Calculations/DemandIntervals.xlsx | Table_ShiftDemandHRSCheck | No |
| No | UNITS/Unit2/1. Input/1-AllocationExtracted.xlsx | AllocationExtractedCheck | No |
| No | UNITS/Unit2/1. Input/Demand-MasterRoster Manual Read.xlsx | LocRoleDayShift%CHECK | No |
| No | UNITS/Unit2/2. Calculations/Demand.xlsx | ShiftDemandHCAverageCheck | No |
| No | UNITS/Unit2/2. Calculations/Demand.xlsx | Table_ShiftDemandUnitINTERVALCheck | No |
| No | UNITS/Unit2/2. Calculations/DemandIntervals.xlsx | ShiftDemandUnitINTERVALCheck | No |
| No | UNITS/Unit2/2. Calculations/DemandIntervals.xlsx | Table_ShiftDemandHRSCheck | No |

## Source-backed files with checks or gates

| Default | Workbook | Named checks | Hard gates | Captures | Empty/count branches | Unreferenced checks |
|---|---|---:|---:|---:|---:|---:|
| Yes | 2. Calculations/Cost/Cost..xlsx | 0 | 4 | 0 | 3 | 0 |
| Yes | 2. Calculations/E-O-I/Effort-All.xlsx | 1 | 5 | 1 | 4 | 0 |
| Yes | 2. Calculations/E-O-I/EffortOutcomes.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | 2. Calculations/E-O-I/Inefficiencies.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | 2. Calculations/E-O-I/StafMasterList-All.xlsx | 0 | 2 | 1 | 1 | 0 |
| Yes | 2. Calculations/EOW/2DRead.xlsx | 1 | 0 | 0 | 0 | 1 |
| Yes | 2. Calculations/EOW/EffortOutcomeLogXY.xlsx | 1 | 0 | 0 | 0 | 1 |
| Yes | 2. Calculations/Tableau Connection.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/BD/2. Calculations/Capacity.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/JH-RY/2. Calculations/Capacity-ShiftAvailability.xlsx | 15 | 19 | 6 | 13 | 0 |
| Yes | UNITS/TE/1. Input/1-AllocationExtracted.xlsx | 1 | 2 | 1 | 1 | 1 |
| Yes | UNITS/TE/1. Input/2-DemandExtract.xlsx | 6 | 21 | 2 | 20 | 0 |
| Yes | UNITS/TE/1. Input/Demand-MasterRoster Manual Read.xlsx | 13 | 9 | 1 | 11 | 1 |
| Yes | UNITS/TE/1. Input/Worker Reconciliation.xlsx | 10 | 8 | 0 | 8 | 0 |
| Yes | UNITS/TE/2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx | 5 | 12 | 4 | 12 | 0 |
| Yes | UNITS/TE/2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx | 1 | 1 | 1 | 1 | 1 |
| Yes | UNITS/TE/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx | 5 | 8 | 2 | 6 | 2 |
| Yes | UNITS/TE/2. Calculations/Allocation.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/TE/2. Calculations/AllocationByShiftAverage.xlsx | 1 | 2 | 2 | 2 | 0 |
| Yes | UNITS/TE/2. Calculations/Capacity-ShiftAvailability.xlsx | 14 | 19 | 6 | 13 | 0 |
| Yes | UNITS/TE/2. Calculations/Capacity.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/TE/2. Calculations/Demand.xlsx | 3 | 1 | 1 | 1 | 2 |
| Yes | UNITS/TE/2. Calculations/DemandIntervals.xlsx | 2 | 1 | 1 | 1 | 2 |
| Yes | UNITS/TE/2. Calculations/Effort.xlsx | 2 | 7 | 4 | 6 | 0 |
| Yes | UNITS/TE/2. Calculations/Intervals.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/TE/2. Calculations/MutliRoleCheck.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/TE/2. Calculations/Settings Data.xlsx | 0 | 2 | 0 | 2 | 0 |
| Yes | UNITS/TE/2. Calculations/Shift-StaffDistribution.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/TE/2. Calculations/Shifts.xlsx | 0 | 1 | 1 | 1 | 0 |
| Yes | UNITS/TE/2. Calculations/StaffListMaster.xlsx | 2 | 9 | 4 | 9 | 0 |
| No | UNITS/Unit1/1. Input/1-AllocationExtracted.xlsx | 1 | 2 | 1 | 1 | 1 |
| No | UNITS/Unit1/1. Input/2-DemandExtract.xlsx | 6 | 18 | 2 | 17 | 0 |
| No | UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx | 13 | 7 | 1 | 10 | 1 |
| No | UNITS/Unit1/2. Calculations/Allocation.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit1/2. Calculations/AllocationByShiftAverage.xlsx | 1 | 2 | 2 | 2 | 0 |
| No | UNITS/Unit1/2. Calculations/Capacity-ShiftAvailability.xlsx | 9 | 9 | 5 | 15 | 0 |
| No | UNITS/Unit1/2. Calculations/Capacity.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit1/2. Calculations/Demand.xlsx | 3 | 1 | 1 | 1 | 2 |
| No | UNITS/Unit1/2. Calculations/DemandIntervals.xlsx | 2 | 1 | 1 | 1 | 2 |
| No | UNITS/Unit1/2. Calculations/Effort.xlsx | 2 | 8 | 4 | 7 | 0 |
| No | UNITS/Unit1/2. Calculations/Intervals.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit1/2. Calculations/MutliRoleCheck.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit1/2. Calculations/Shift-StaffDistribution.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit1/2. Calculations/Shifts.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit1/2. Calculations/StaffListMaster.xlsx | 2 | 9 | 4 | 9 | 0 |
| No | UNITS/Unit2/1. Input/1-AllocationExtracted.xlsx | 1 | 2 | 1 | 1 | 1 |
| No | UNITS/Unit2/1. Input/2-DemandExtract.xlsx | 4 | 10 | 2 | 9 | 0 |
| No | UNITS/Unit2/1. Input/Demand-MasterRoster Manual Read.xlsx | 13 | 3 | 0 | 7 | 1 |
| No | UNITS/Unit2/2. Calculations/Allocation.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit2/2. Calculations/AllocationByShiftAverage.xlsx | 1 | 2 | 2 | 2 | 0 |
| No | UNITS/Unit2/2. Calculations/Capacity-ShiftAvailability.xlsx | 7 | 8 | 5 | 11 | 0 |
| No | UNITS/Unit2/2. Calculations/Capacity.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit2/2. Calculations/Demand.xlsx | 3 | 1 | 1 | 1 | 2 |
| No | UNITS/Unit2/2. Calculations/DemandIntervals.xlsx | 2 | 1 | 1 | 1 | 2 |
| No | UNITS/Unit2/2. Calculations/Effort.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit2/2. Calculations/Intervals.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit2/2. Calculations/MutliRoleCheck.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit2/2. Calculations/Shift-StaffDistribution.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit2/2. Calculations/Shifts.xlsx | 0 | 1 | 1 | 1 | 0 |
| No | UNITS/Unit2/2. Calculations/StaffListMaster.xlsx | 0 | 1 | 1 | 1 | 0 |

## Detailed inventories

- `workbook-check-inventory.csv` contains every catalogue workbook, including missing-source blockers.
- `query-check-inventory.csv` contains every query with a check-like name, explicit M error, error capture/fallback, or empty/count branch.

Counts are static-source indicators, not proof that a condition fails at runtime. Workbook load settings and current query results remain unverified unless they are exposed through an approved source and a runner-readable contract.
