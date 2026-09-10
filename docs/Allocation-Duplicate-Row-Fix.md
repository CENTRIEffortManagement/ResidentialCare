# Allocation duplicate-row source fix

## Approved source and scope

The user approved applying the reviewed patch to the exact adjacent source:

`CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AllocationByShiftAverage.xlsx_PowerQuery.m`

This is a narrow, user-approved exception to the normal `Workflows/` source location. The user also approved verifying this source against the exact workbook as the source-of-truth process for this edit. Before editing, all section definitions matched the embedded workbook definitions after normalizing line endings and blank lines. Both file hashes were unchanged from that verification when the patch was applied.

- Workbook: `CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AllocationByShiftAverage.xlsx`
- Workbook SHA-256 before the source edit: `575bee122bc59da6a8fc7064f6c94b002d72528048a97405d41fd45262ad2881`
- Approved M source SHA-256 before the edit: `fd3953ad731defd6810699e2a918a33347d9794e1f11ca5d9c2845713e650c83`

## Resulting behavior

`RoleShiftAllocation_StartEffort` selects matching-role `IntervalStart` records and aggregates effective effort by shift date, role and shift period. `RoleShiftAllocation_Calculated` matches durations on both role and shift period, retaining a match count instead of expanding duplicate matches.

`RoleShiftAllocation_CHECK` exposes failures for duplicate or missing keys, non-start endpoints, ambiguous/missing/nonpositive/non-finite durations, invalid effort or FTE, missing start keys, and FTE-to-hours reconciliation. Both output queries block publication when any required check fails.

`RoleShiftIntervalAllocation` now uses the same validated calculation as `RoleShiftAllocation`. Its query name and seven-column output interface remain intact so the existing `Table_RoleShiftIntervalAllocation` load can continue feeding Allocation. Effort remains downstream of that interface.

## Validation and deployment state

The reviewed source passed the installed Microsoft Power Query parser. Regression reconstruction using the saved 33,654 resource rows produced 252 calendar rows, including nine zero-allocation padding rows. The Unit1 / AIN / 20 July 2026 / AM example produced 10.836616161616162 FTE. Missing, duplicate, zero, negative, null, infinite and NaN duration cases were rejected; endpoint-boundary and role-expansion cases passed.

The regression reconstruction is Python, not execution of the M queries in Excel. The final source is syntax-checked after application. Query headers, named steps, dependencies, final `in` targets and existing output interfaces were reviewed.

This approval covers the M source edit only. No workbook synchronization, refresh, save or close is included. A later explicitly authorized sync must preserve the existing output-table binding and keep the new staging/check queries connection-only. After an authorized refresh, verify 252 rows at the expected calendar grain in Allocation and Effort, passing checks, the example FTE and Tableau's extract.
