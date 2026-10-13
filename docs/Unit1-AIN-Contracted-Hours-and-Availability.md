# Unit1/AIN contracted hours and availability — handover note

Status: implementation snapshot as of 19 September 2026. This note describes the approved Power Query source and the checks reported during the Unit1/AIN rollout. It is not proof that every workbook still matches its adjacent source or that the complete capacity chain has since been refreshed.

Documentation review: 2026-10-05. Historical rollout observations remain dated below. See [Capacity determination notes](Capacity-Determination-Notes.md) for the current BD/AIN explanation, review findings, and proposed improvements. The review updated documentation only; it did not change Power Query or workbooks.

## Scope and core distinction

This change calculates employee contract cap targets and reductions for *redistributable availability* for Unit1/AIN resources. The distribution stages apply selected reductions; a calculated cap is not proof that the published total reaches that limit. It does **not** trim existing roster allocations to contract. An employee may therefore be allocated above their agreed contract; a separate over-contract output reports that condition. Contract limits are at employee/Resource and whole-roster grain, not Resource–Period or date/shift grain.

The 19 September rollout changed only Unit1/AIN. Do not copy this code to RN, AINC4, Unit2, or another client without first reviewing folder-specific filters, role decisions, paths, and outputs. Unit1/AIN was the working reference; those other roles/units were behind it at that time. A.2 was deliberately left functionally unchanged in that rollout; its current source coverage is qualified below.

The contract source is the existing Unit1 `1. Input/Worker's Report.xlsx`, `Table1`, fields `Employee Code` and `Contracted FN Hours` (hours per fortnight). No new input workbook or Worker Reconciliation calculation was introduced. Worker Reconciliation remains authoritative for employee identity and preferred role. In particular, a preferred-AINC4 worker does not enter AIN availability and their contract is not proportionally divided into AIN capacity.

## Source and calculation lineage

All Unit1 paths below are relative to `CLIENT/DATExx-Whiddon/UNITS/Unit1/`.

| Stage | Adjacent Power Query source | Contract-relevant interface |
| --- | --- | --- |
| Availability/reconciliation | `2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m` | `IMPORT Worker's Report` → `WorkerContracts_Prepare` → `ReconciledWorkers_Prepare`/`ReconciledWorkers_CHECK` → loaded `Availability-StaffList` |
| Resource master | `2. Calculations/StaffListMaster.xlsx_PowerQuery.m` | `EXTRACT Availability Staff List` reads the `Availability-StaffList` **sheet**; `ResourceContract` and `ResourceContract_CHECK` are staging/check queries; loaded `Table_Masterlist` carries the cap |
| First AIN distribution | `2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx_PowerQuery.m` | `A1ResourceContract_CHECK`; Resource-grain join in `ResAv-C'` |
| Second AIN distribution | `2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx_PowerQuery.m` | No contract edit or new join; uses A.1's calculated reduction |
| Final AIN distribution | `2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx_PowerQuery.m` | `BResourceContract_CHECK`; Resource-grain caps in `ResRosterAvailabilityCapReduction`, `ResRosterAvailabilityC##CapReduction`, `ResMaxAvailability` |
| Unit1 effort | `2. Calculations/Effort.xlsx_PowerQuery.m` | `ResourceContractOverage_Prepare`/`ResourceContractOverage_CHECK` → separately loaded `ResourceContractOverage` |
| Organisation effort | `CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All.xlsx_PowerQuery.m` | Facility1 Effort import → `EXTRACT Facility1 ResourceContractOverage` → check → separately loaded `ResourceContractOverage` |

The `Availability-StaffList` extension carries reconciled `EmployeeID`, facility, employment type, preferred role, and fortnightly contract hours. Employee codes and IDs are matched as text. Worker's Report contract preparation accepts repeated *identical* contract rows and records conflicting, negative, non-numeric, or non-finite values as issues. Relevant availability contract issues block publication; unusable contracts for allocation-only resources remain diagnostic warnings. A missing or zero non-casual contract for an availability resource is a blocking check failure. The preferred-role decision remains the one made in reconciliation.

`Table_Masterlist` has one row per Resource and retains allocation-only resources as well as availability resources. Its contract-related columns include `EmployeeID`, `Facility-Abbrev`, `EmploymentType`, `PreferredRole`, `Worker Record Status`, `Worker Contract Issue`, `Contracted FN Hours`, `Roster Start`, `Roster End`, `Roster Fortnights`, `Contracted Roster Hours`, `Contracted Shift Equivalent`, `Contracted Shifts`, `Settings Max Availability`, `Effective Shift Cap`, and `Limit Basis`. This loaded Excel table is the interface consumed by A.1, B, and Effort. Do not join it at Resource–Period grain: that would repeat roster totals and can multiply rows.

## Cap arithmetic and exceptions

The master list derives distinct, contiguous Settings roster dates from `Settings Data.xlsx/PermutationDimensions`; their count must be a positive multiple of 14. `Roster Fortnights = distinct roster dates ÷ 14`. The 19 September snapshot used a 28-day roster, giving 2 fortnights, with `ShiftDuration` of 7.6 hours and Settings `MaxAvailability` of 20 shifts. These are the values used in the examples below. The source reads them from Settings; they are not fixed code constants or verified current input values.

```text
Contracted Roster Hours      = Contracted FN Hours × Roster Fortnights
Contracted Shift Equivalent  = Contracted Roster Hours ÷ ShiftDuration
Contracted Shifts            = floor(Contracted Shift Equivalent)
Effective Shift Cap          = min(Settings Max Availability, Contracted Shifts)
```

The floor applies to *capacity only*. For example, 76 fortnightly hours over 28 days at 7.6 hours/shift gives 152 hours and 20 shifts; 38 fortnightly hours gives 76 hours and a 10-shift cap. A larger contract has a calculated cap of 20 under those example Settings values. Exact decimal hours, not floored shifts, determine whether allocation exceeds contract.

There are deliberate exceptions and diagnostics:

| Resource case | Master-list treatment | Downstream consequence |
| --- | --- | --- |
| Availability, positive valid contract | Contract-derived shift cap, bounded by Settings maximum | Must have one valid cap for positive AIN availability |
| Availability, casual, blank or zero contract | Settings cap; `Limit Basis = Casual fallback - Settings maximum` | Can contribute capacity; no exact contract exists for an over-contract finding, so it is diagnostic only |
| Availability, non-casual, blank or zero contract | `ResourceContract_CHECK` failure | Must be corrected before publication; do not silently replace with zero or Settings maximum |
| Allocation-only, found in Worker’s Report | May retain a null `Effective Shift Cap`; `Limit Basis = Allocation-only resource` | Does not require availability coverage merely because an allocation exists |
| Allocation-only, found with an unusable contract | `Worker Record Status = Found - <issue>`, null contract hours and cap, and a check warning | Remains visible for correction; does not block the master-list output merely because its contract is unusable |
| Allocation-only, absent from Worker’s Report | `Worker Record Status = Not found`, Settings-maximum fallback, and a check warning | Flags the missing record without fabricating contract hours or an over-contract result; B requires zero AIN availability for this fallback |

The missing-worker fallback arose because some Resources appeared in allocation but not Worker’s Report. It must remain visible through `Worker Record Status`/the check, rather than being mistaken for a genuine 0-hour contract. `ResourceContract_CHECK` is the place to inspect non-casual zero/missing contracts and missing-worker diagnostics. Warnings and blocking failures are different: output queries gate on `Status = Fail`.

## Where the capacity limit is applied

A.1 imports `Table_Masterlist` and joins its one cap row per Resource inside `ResAv-C'`, after roster availability has been aggregated to Resource grain. It calculates a capped roster total and the required resource reduction; subsequent stages apply selected reductions to availability rows. `A1ResourceContract_CHECK` requires complete, unique, positive, whole-shift cap coverage for AIN availability contributors and validates their preferred role. A.1 also contains a local repair where the adjacent Settings input `Shifts` is expanded to the downstream `Shift` field. This is not a mandate to rename fields across other roles.

In the rollout, A.2 used the A.1 reduction as before. B imports the same loaded `Table_Masterlist` and calculates the per-Resource cap and required reductions in its three cap/reduction queries named above. Its distribution queries select the reductions applied to individual rows. `BResourceContract_CHECK` checks uniqueness, cap validity, and coverage of Resources with *positive* AIN availability; passing these input checks alone does not establish that the final count reaches its cap. A Resource with zero availability should pass through as zero; a positive-availability Resource with a missing cap must fail visibly. The earlier null-to-zero workaround should not be reinstated as a general substitute for contract validation. There is no contract-hours join at Resource–Period grain.

Raw `ResAvailability` can exceed 20 before the relevant capping step. A preview of that upstream field is not by itself evidence that the final redistributable availability exceeded its cap. Inspect the capped/reduction queries and final outputs, with the checks, when assessing a breach.

## Over-contract reporting and Effort-All

Unit1 Effort uses `1-AllocationExtracted.xlsx/AllocationExtracted` actual `Hours`, restricts rows to Settings roster dates, normalizes employee `Code` as text, and sums hours **once per employee across all allocated roles**. It attaches the preferred AIN Resource from `Table_Masterlist`, then compares `Allocated Roster Hours` with exact `Contracted Roster Hours`. AIN and alternative-role allocations for the same employee are not separate contract allowances.

`Over-Contract Hours = Allocated Roster Hours − Contracted Roster Hours`; shift equivalents divide the corresponding exact hours by `ShiftDuration`. `Is Over Contract` requires a positive difference greater than the numeric tolerance `0.00000001`. Equality is not a breach. Casual fallbacks have no exact positive contract and are excluded from breach publication. The loaded `ResourceContractOverage` table contains **only positive overages**, one employee/preferred Resource/roster, with facility, preferred role, employee/name/employment type, roster dates/fortnights, contract and allocation hours/equivalents, cap context, allocated-role list, `Limit Basis`, `Is Over Contract`, and `Coverage Scope = Unit1/AIN`. It is not a list of all employees or all allocated roles.

`ResourceContractOverage_CHECK` tests identity/contract coverage, roster alignment, valid allocation rows, exact arithmetic, unique output keys, positive published rows, and Unit1/AIN scope; casual fallbacks are diagnostic warnings. The user reported that this check had no failures and that Unit1 Effort refreshed. Effort's `ResShiftAllocation` must remain a loaded Excel table because Effort-All imports it; the earlier missing-table error was an interface/load issue, not a reason to remove that query.

The 19 September rollout described Effort-All extracting the new table from its existing Facility1 Effort import, checking that Facility1 resolves to Unit1, preserving the explicit Unit1/AIN scope, and publishing a separate `ResourceContractOverage` table. Do **not** append contract roster totals to `Availability Effort` or `EffortAllMatrixAG1_1D`: those are date/shift-grain facts and would duplicate totals. At that time, the user reported that Effort-All refreshed without incident and that the Unit1 Effort and Effort-All overage tables had matching coverage. The current source has changed; this historical report does not establish its current publication behaviour.

The availability route into Effort-All is separate: each configured facility's Effort `RoleShiftAvailabilities` → its `EXTRACT FacilityN RoleShiftAvailabilities` → `Availabilities Appended` → `Availability Effort` → availability-developed matrices. Effort-All does **not** import Effort's compatibility query `AvailabilityDeveloped2`. The 19 September snapshot recorded a pre-existing `Resource = 108` filter. The source reviewed on 5 October appends Facilities 1, 2 and 3 and no longer contains that filter; the historical observation is not a description of its current scope.

## Documentation review and current applicability

The 5 October read-only review found matching target headers for the Unit1 Availability, StaffListMaster, AIN A.1, AIN B, and Effort sources, and for organisation Effort-All. The contract arithmetic, Resource grain, core exception policy, and Unit1 Effort exact-hours overage route agree with those sources. Matching headers do not establish current workbook parity or loaded-table contents.

The adjacent Unit1/AIN A.2 source header identifies a legacy `CLIENT/DATExx/UNITS/Unit1/2. Calculations/RoleA/` workbook. Its current Unit1/AIN behaviour was not verified from that file. The Unit1 Availability source also explicitly scopes reconciled workers and relevant contract issues to facility abbreviation `BD`; applying this note requires confirming that Unit1 is intended to represent that facility. This is a scope prerequisite, not evidence that the unit folder is wrong.

In the 5 October Effort-All source, `ResourceContractOverage` returns the rows of `ResourceContractOverage_CHECK`, rather than the employee overage records described in the September rollout. Its availability scope has also changed as noted above. Reported historical refreshes, source header dates, and this documentation review do not verify current Settings values, workbook load configuration, saved output data, synchronization, or a complete refresh. The current BD/AIN review and proposed improvements are maintained in [Capacity determination notes](Capacity-Determination-Notes.md).

## Refresh, validation, and future work

The intended dependency order is `Capacity-ShiftAvailability` → `StaffListMaster` → A.1 → A.2 → B → downstream capacity/Effort → Effort-All → outcomes/efficiencies/Tableau connection as appropriate. For the contract overage specifically, the loaded `Availability-StaffList`, `Table_Masterlist`, and Unit1 Effort `ResourceContractOverage` must exist before downstream workbooks read them. `ResourceContract` and the named checks are connection-only staging/check queries; do not create duplicate sheet loads just because a query exists.

The implementation was committed and pushed in `96f42f4` (`Add Unit1 AIN contract availability caps and overage reporting`). It includes the six adjacent source/workbook pairs in the table above other than unchanged A.2. The later `e150070` commit recorded organisation availability refresh timelines. These commit references identify a snapshot, not a guarantee that the current working tree or live Excel files are unchanged.

The user performed targeted workbook refresh/checks through Effort-All and reported no overage-check failures. Subsequent batch runs refreshed the organisation O1/O2/O5 stages and saved/closed successfully; O3 cost and O4 history were skipped and existing outputs consumed. Those runs do **not** establish an end-to-end A.1 → A.2 → B capacity result. A complete capacity refresh and business reconciliation remain explicit future work. Exact post-change source-to-workbook M parity was not independently audited in this handover; do not state that it was.

Before further changes, inspect the relevant adjacent `.m` source and its query interfaces, then follow the repo's approved source-of-truth and path conventions. Per `AGENTS.md`, do not inspect Excel workbooks without explicit instruction, and do not sync an M source into a workbook merely because it changed. A named workbook/source sync requires explicit opt-in, a closed-workbook check, M validation, recoverable backup, approved sync method, and post-sync re-extraction comparison. Sync does not imply refresh. For authorised runner work, follow `docs/ResidentialCare-Runner-Agent-Instructions.md`.
