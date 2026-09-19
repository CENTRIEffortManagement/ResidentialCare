# Four-worker source implementation handoff

## Authority and scope

The user confirmed `.m` has higher precedence than `.xlsx`. The exact adjacent Unit1 Availability and AIN B `.m` files are the authoritative sources for this bounded repair, as a narrow exception to the usual `Workflows/` location. Read-only source-gate snapshots and hashes are in `availability.source-gate.json` and `ain-b.source-gate.json`. Availability's embedded definition matched its prior `.m`; B's embedded definition differed from its prior `.m`. B's `.m` was edited without importing the workbook definition. The historical revised Availability M is retained only as `availability.revised-reference.m` for comparison.

At the initial bounded-test handoff, the four-worker filters remained in the Availability source. The Worker's Report filter included 19822, 19835, 24562 and 25376 once each. No source for AIN A.1, A.2, StaffListMaster, or Capacity was edited. Later Excel saves and the full-population source edit are recorded below.

## Rule and output

Availability separates AVAIL, UNAVAIL, recognised leave and unknown reasons. AVAIL resets by calendar date; UNAVAIL subtracts only recorded shift overlap; distinct recognised leave of at least `ShiftDuration` blocks every shift starting on that date. `AvailabilityRules_TEST` and `ResDayShift_CHECK` gate publication, while `AvailabilityReasons_DIAGNOSTICS` surfaces unknown reasons. Existing contract and staff/output columns are retained.

AIN B exposes `CapacityDistribB_AVAILABILITY_LINEAGE_CHECK`. Every positive A.2 C# Role/Resource/Period must have exactly one matching positive original A.1 Availability row. `C###TABLE B` fails when that check returns rows. Its C## and C### arithmetic remains unchanged. Today's saved A.2 has 81 positive C# rows across Resources 10, 12, 26 and 27, while the later saved A.1 has no original availability for those four. This is a mixed saved snapshot, not evidence that an ordered rerun will still fail. The check should reject that mismatch if presented with it, and should pass when A.1, A.2 and B use consistent inputs.

`day_shift_matrix_before.csv` is the saved old-rule snapshot. `day_shift_matrix.csv` and `day_shift_matrix_expected.csv` have 336 rows and predict 108 positive shifts under the revised rule, compared with zero in the saved `ResDayShift`. Expected rows are independent calculations from the frozen source-record evidence; they are not refreshed M output.

## External file access audit

| Source | External file | Named query and root |
| --- | --- | --- |
| Availability | `1. Input/1-AllocationExtracted.xlsx` | `IMPORT MultiRolesResRolesProportion` via `FilePath - 1Input` |
| Availability | `1. Input/Worker Reconciliation.xlsx` | `IMPORT Reconciled Workers` via `FilePath - 1Input` |
| Availability | `1. Input/Worker's Report.xlsx` | `IMPORT Worker's Report` via `FilePath - 1Input` |
| Availability | `1. Input/whiddon_availability_leave_extraction.xlsx` | `IMPORT Availability Leave Source` via `FilePath - 1Input` |
| Availability | `2. Calculations/Settings Data.xlsx` | `IMPORT Availability Settings` via `FilePath - 2Calculations` |
| AIN B | `2. Calculations/Settings Data.xlsx` | `IMPORTSource Settings Data` via parent of `RolePath` |
| AIN B | `2. Calculations/StaffListMaster.xlsx` | `IMPORTSource StaffListMaster` via parent of `RolePath` |
| AIN B | `2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx` | `IMPORTSource A1` via `RolePath` |
| AIN B | `2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx` | `IMPORT ResPeriodAvailabilityCapped(C#)` via `RolePath` |
| Both | Public `CentriSyncPaths.xlsx` | Sole fixed-location bootstrap inside the authoritative path resolver |

Both resolvers read `FilePathUrl` as a one-row `FilePath` table or as the existing named cell exposed by Excel as `Column1`. They normalize the latter to `FilePath`, then normalize local/SharePoint and CELL filename forms, choose a case-insensitive boundary-safe longest mapping, and build imports from the resolved workbook folder. No legacy `Folder` input or second resolver was retained in the edited sources.

## Validation and remaining workbook prerequisite

`validate_source_edits.ps1` passed: 44 unique Availability and 74 unique B shared definitions, balanced lexical delimiters, revised-rule presence, old-rule absence and B publication gate. `git diff --check` passed. The 336-row matrix was regenerated, with 108 positive expected rows; four 20 July spot cases passed. All 16 source records in the sample's prior `Leave/other exclusion` category were checked against the saved raw reason labels and match the recognised leave list. Runtime Power Query evaluation was not performed; the available source validator does not parse or execute M.

Read-only inspection found a `FilePathUrl` **named cell** in each workbook (`UnitL1RolePath!$B$1` for Availability and `UnitL1Path!$B$1` for B), with no `FilePathUrl` table. The Power Query resolvers now accept those existing cells. Source synchronization and refresh remain separate steps. After source sync, re-extract the embedded definitions to compare with the authoritative `.m`. Refresh Availability, StaffListMaster, AIN A.1, A.2 and B in order, including runner-declared prerequisites. A.1 reads both Availability and StaffListMaster. Regenerate the actual four-worker matrix and compare the four IDs through every stage before deciding whether the source repair fixed Capacity. Refresh Capacity too if judging its published workbook output. A.1/A.2/B alone cannot test the revised availability rule while Availability still has its old saved result.

## Later Excel error, 19 September 2026

Excel displayed `[Expression.Error] FilePathUrl must have one row and a FilePath column.` That wording matches the former Availability resolver and occurs before worker-shift calculations. The Availability workbook was saved at 16:17 and was open/locked during the follow-up check, so its embedded M and current worksheet structure were not re-extracted then. The B workbook was saved at 16:16; a read-only package check still found `FilePathUrl` as a named cell and no table. These later saves supersede the earlier workbook hashes as evidence of current file state. Do not interpret this error as a failed availability rule or a completed A.1/A.2/B test. The authoritative `.m` path input now accepts the existing named cells; verify embedded definitions after source sync, then perform the ordered refresh.

The user then explicitly requested the path fix in `Capacity-ShiftAvailability.xlsx`. Once that workbook was closed, a new read-only source gate found its embedded M matched the then-current authoritative `.m` (workbook SHA-256 `AA4094B572F5488F2E2678F5C55139D09C923AD38A9383DC8B2DB519B8DEE1E2`). `UnitL1RolePath!B1` contains `=CELL("filename")` and currently resolves to this workbook. `FilePathUrl` is a defined name pointing at B1; there is no table by that name. An attempted native Excel automation launch was rejected by execution approval (`rejected by user`); no workbook change or refresh was made. The subsequent user correction clarified that the fix belongs in Power Query. The two authoritative `.m` resolvers were updated to accept the existing named cells. Their workbook definitions have not been synchronized or refreshed in this step.

## Saved Availability refresh and walkthrough

The user synchronized and refreshed `Capacity-ShiftAvailability.xlsx` and saved it at 16:32 on 19 September. A read-only post-refresh source gate confirms that its embedded M matches the adjacent authoritative `.m` (workbook SHA-256 `8989706BB54804A412C7969E9165CD81DDA47ED95A29214F00ADB47429566846`). `compare_refreshed_availability.ps1` checked all 336 `WorkerShiftSegments` rows against the independent expected matrix and all 108 published `ResDayShift` rows: zero mismatches, with all 12 saved `ResDayShift_CHECK` rows passing. The published counts by worker are 19822: 25, 19835: 5, 24562: 32 and 25376: 46. `day_shift_matrix_reconciled.csv` contains the actual result beside source records and expected values; `availability-walkthrough.md` explains the derivation. AIN A.1, A.2, B and Capacity have not been rerun as part of this verification.

## Full-population source edit

After the user reviewed the four-worker result, the three sample filters were removed from the authoritative Availability `.m`: `IMPORT Reconciled Workers`, `IMPORT Worker's Report` and `IMPORT Availability Leave Source` now pass their complete selected tables to the existing downstream preparation. The BD facility restriction in the availability source and the existing nursing-role and identity checks remain. Immediately before this edit, a read-only source gate on the workbook saved at 16:48 found its embedded M still matched the adjacent `.m` (workbook SHA-256 `48CBC1695B8697FFD8D95F88689AC0E792ADBA79E1B56C3E45A3374D707397B3`). Source validation passed after the edit; the workbook has not been synchronized or refreshed with the full-population source. The four-worker matrix remains a frozen test record.

The user subsequently reported running Unit1 and is checking the results. The generated actual-run note for `20260919-165249-04a88aab` reports **Completed** technical refresh and explicitly says business reconciliation was not evaluated. This note has not independently verified that run's embedded M, full BD `ResDayShift`, AIN A.1/A.2/B lineage or Capacity. The preceding synchronization statement describes the state at the time of the source edit, before the user-reported run.
