# BD Balance removal: manual synchronization and runtime test

Status: source comparison and M syntax validation passed. Synthetic native evaluation could not complete in the sandbox: Windows profile encryption failed, and the requested run outside the sandbox was cancelled. No native fixture pass is claimed. Manual workbook synchronization and runtime testing have **not** been performed by this change.

Manual observations reported by the user on 25 October 2026: Capacity checks were reported as okay. The supplied Staff `ResourceContract_CHECK` result contains ten blocking checks with `Pass` and zero failures, plus one `Workers using Settings roster-cap fallback` warning for one worker. That worker's exception still requires review before final acceptance. These observations were not independently rerun or tied to refreshed workbook hashes; the initial source-verification records below remain historical baseline evidence.

## Exact targets

| Workbook | Approved M source |
|---|---|
| `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx` | `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m` |
| `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/StaffListMaster.xlsx` | `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/StaffListMaster.xlsx_PowerQuery.m` |

Worker Reconciliation and all other units are outside this change. Keep their files unchanged. The approved workbook-derived baseline contains 71 Availability queries and 27 Staff queries. Use that baseline and the final source-difference evidence in this folder; the older adjacent Availability source was not a reliable complete workbook baseline.

Original Balance-based M archives are recorded in `outputs/bd-balance-removal-20261025/archive-manifest.json` under `archives/BD/`, with the suffix `.xlsx_PowerQuery-BALANCE.m`. The earlier adjacent Availability source is separately retained at `outputs/bd-balance-removal-20261025/pre-baseline/BD/Capacity-ShiftAvailability.xlsx_PowerQuery-pre-baseline.m`; it is evidence, not a sync target.

## Approved source changes and checks

The exact query comparison is in `outputs/bd-balance-removal-20261025/source-validation.json`. Availability changes from **71 to 76 queries**:

- Add six connection-only inputs/preparation queries: `IMPORT Published Roster`, `PublishedRoster_Prepare`, `RosterWorkers_Prepare`, `WorkerEmployment_Prepare`, `WorkerEmployment_Selected` and `AvailabilityMembership_Prepare`.
- Remove only `IMPORT Reconciled Workers`.
- Update `IMPORT Worker's Report`, `ReconciledWorkerMembership_Prepare`, `WorkerContracts_Prepare`, `ReconciledWorkers_Prepare` and `ReconciledWorkers_CHECK`; `ReconciledWorkers_Eligible` has a comment-only change.

Staff remains **27 queries**, byte-for-byte unchanged from its approved archive, so it needs no formula synchronization. Published output interfaces, the existing calculation definitions and facility mapping are unchanged. The source comparison verified the archives, source boundary and unchanged workbook hashes; it reports the Balance workbook dependency removed from the revised Availability source.

Both named M sources passed the installed Microsoft Power Query extension parser, with source hashes recorded in `outputs/bd-balance-removal-20261025/verification/syntax-validation.json`. This is a source parse result, not a workbook evaluation.

Memory/performance source review found narrow scalar-column buffers and a unique worker-employment location lookup to limit join expansion. Separate checks can reevaluate source reads; a shared query or buffer is not a durable cross-query cache. Runtime and memory use have not been measured.

The isolated fixture runner substitutes synthetic tables for every external input and rejects external connector calls. Before manual synchronization, run it from the repository root in a normal PowerShell 7 session with the installed Power Query SDK available:

```powershell
pwsh -NoProfile -File outputs/bd-balance-removal-20261025/tests/run-native-fixtures.ps1
```

Inspect `tests/summary.json` and `tests/native-result.json` after a completed run. Require `Status = "Passed"`, zero failed rows and zero external data sources before counting these fixtures as passed. The script rebuilds its generated test extension from the current original M sources; do not run an earlier generated `.mez` directly. `-PrepareOnly` builds the isolated fixtures without launching the SDK and does not establish a native test pass.

## Manual synchronization

1. Confirm the two exact workbooks are closed and that neither workbook nor approved M source has changed since the recorded baseline/source checks. Keep the recorded source archives; retain a recoverable byte-for-byte workbook copy before any authorized workbook mutation. Stop and resolve any version difference before applying formulas.
2. Apply only the approved query differences to the named workbook/source pair through Excel's native Power Query interface or a separately authorized native-query sync. Preserve existing worksheets, tables, output bindings, connections and unrelated query definitions. Retain the current load settings of existing outputs. Keep newly added preparation/diagnostic queries connection-only unless the approved query manifest specifies a load.
3. Apply the exact Availability query changes listed above. Remove only the obsolete `IMPORT Reconciled Workers` query after its consumers have been updated. Check dependencies and any load binding before deleting it; do not recreate the workbook or edit its XLSX package. Staff needs no formula synchronization; its refresh remains part of the runtime test.
4. Do not run the historical generic sync helper unchanged. It rejects workbook-only queries, and the newer Staff helper requires unchanged query names/counts. This change can add/remove queries and needs that exact difference handled deliberately.
5. Save and close Availability after the authorized synchronization. Re-extract its Power Query through the approved mechanism and compare all 76 query names and formulas with its approved M source, allowing only the established line-ending normalization. Record the exact workbook/source hashes, query changes, comparison result and save/closed status. Keep Staff unchanged until its separately authorized refresh. Synchronization alone does not refresh.

## Runtime test sequence

Run this sequence only after an explicit instruction to refresh the named BD workbooks. For runner execution, follow `docs/ResidentialCare-Runner-Agent-Instructions.md` and include its declared prerequisites; a source test or preview is not a workbook refresh.

1. Refresh **Capacity-ShiftAvailability.xlsx first** from the intended saved BD inputs. Inspect its evaluated checks and outputs below, then accept and save this result. Record its saved version and input versions.
2. Refresh **StaffListMaster.xlsx only after** the accepted Availability result has been saved. It reads the saved `Availability-StaffList` sheet. Record that exact input version, inspect its evaluated checks and output, then save the accepted result.
3. Stop at the first blocking error. Do not judge downstream role-capacity or Capacity results from earlier saved versions; those workbooks are outside this manual test.

| Workbook | Required evaluated result |
|---|---|
| Availability | Every row in `ReconciledWorkers_CHECK`, `AvailabilityRules_TEST`, `ResDayShift_CHECK` and `AllocationRoleShares_CHECK` has `Status = "Pass"` and `Failures = 0`. |
| Staff | Every blocking row in `ResourceContract_CHECK` has `Status = "Pass"` and `Failures = 0`. The `Workers using Settings roster-cap fallback` row may report `Warning`; review each affected worker in Availability's `CasualRosterShiftCapExceptions` and record the disposition before accepting the result. |

Check queries can be connection-only. Inspect their existing sheet/table when loaded; otherwise inspect the evaluated Power Query result. Do not create new worksheet loads for these tests. `AvailabilityRecords_Invalid` must be empty before accepting publication.

## Visual acceptance and evidence

- Inspect Availability's `Availability-StaffList` and `ResDayShift`, and Staff's `Masterlist` query and existing `Table_Masterlist` load. Confirm headers, row grain, table binding and complete saved output. Membership must come from the intended BD sources, without workers introduced solely through Balance or another facility.
- Confirm an otherwise eligible BD worker remains on the staff list when unavailable for every roster shift. Such a worker can have no positive published `ResDayShift` rows; availability must not become a staff-membership filter.
- Review `WorkerIdentityFailures_DIAGNOSTICS`, `WorkerRolePipeline_DIAGNOSTICS`, `WorkerAvailability_DIAGNOSTICS`, `WorkerShiftRoleCoverage_DIAGNOSTICS` and `CasualRosterShiftCapExceptions` where present. Explain any identity, role, availability or cap exception affecting a representative worker. `WorkerRoleMappings_DIAGNOSTICS` can contain unused mapping rows without a publication failure; distinguish unused rows from mappings affecting retained workers. Staff's `Misaligned Join` and `Misaligned Allocation Availability Names` are diagnostics and cannot add people back into `Masterlist`.
- Choose representative actual EmployeeIDs during the manual review: a BD worker present in both source populations, an allocation-only worker, an availability/leave-only worker, a worker unavailable throughout the roster, a worker with multiple mapped roles, and a casual using the roster-cap fallback. Use a case only when it exists in the actual inputs; record absent cases rather than inventing IDs or counts. Confirm an other-facility/Balance-only example is excluded when such a case exists.
- Compare workers by EmployeeID and the applicable facility key. `Resource` numbers may change when the membership list changes; do not treat renumbering as an employee identity failure.
- Record actual input versions, evaluated check results, output counts, representative IDs and observations, exceptions accepted/rejected, final workbook hashes and saved/closed status. Mark runtime acceptance pending until this evidence exists. Source checks alone do not establish workbook runtime correctness.
