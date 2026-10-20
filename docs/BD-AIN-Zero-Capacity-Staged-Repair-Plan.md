# BD AIN capacity repair and test plan

Prepared: 21 October 2026. Status: Change 1 source implemented; isolated native test pending an engine dependency.

Repair one input interface, test that change, and review the result before starting the next change. Restore current Availability and StaffListMaster outputs before testing the AIN calculation chain. A completed Excel refresh is insufficient evidence: each test must inspect its actual query result, checks and roster dates.

## Scope and baseline

All workbook paths below are relative to `CLIENT/DATExx-Whiddon/UNITS/BD/`.

The read-only diagnosis found these saved results:

| Area | Baseline |
| --- | --- |
| Worker Reconciliation | Alexis Griffin, Ainura Borbueva, Deeanne Tracey and Shelby Bon are present with the role Assistant in Nursing. |
| Availability membership input | The older imported population has Alexis under Care Services Employee and omits the other three workers. Current checks report missing worker-membership columns. |
| Availability records input | Two saved checks report that `Payroll Code` cannot be found. The replacement source header or object has not been established. |
| StaffListMaster | Eight saved checks report `Expected exactly one Table1 table in Worker's Report.xlsx.` The actual replacement source object has not been established. |
| Availability and staff contracts | Saved dates cover 20 July to 16 August 2026. |
| AIN A.1 | Demand covers 28 September to 25 October 2026. All 4,190 imported availability rows have no Settings period match, affecting 83 Resources. Published availability is empty. |
| Allocation identity | Four unmatched Name/Role pairs account for 104 skipped allocation rows. Other Resources continue. |

The four inspected workbook definitions matched their exact extracted sources during diagnosis. Evidence is retained in [the diagnostic profiles](../outputs/bd-ain-capacity-diagnosis-20261021/profiles.json) and [workbook inventory](../outputs/bd-ain-capacity-diagnosis-20261021/inventory.json). These are baseline evidence, never replacement source or workbook targets.

The target roster window must be reconfirmed from the approved current Settings before testing. The September/October dates above describe the observed baseline, not a new hard-coded setting.

## Control of each change

Change 1 source editing has been requested. Source edits, workbook synchronization and refresh remain separate operations under [AGENTS.md](../AGENTS.md). Each implementation instruction must identify its exact source and workbook and state which operations are authorized. Stop after each test and report the result; a passing result does not start the next change automatically.

Use the exact existing extracted `.m` named in the implementation instruction and repeat the approved source-of-truth check before editing. Preserve existing targets; do not substitute diagnostic embedded copies, backups or another unit's source. Use `pq.project.json` for project paths and the `Workflows/` convention for any new product source. Source relocation is outside this repair.

For an authorized sync, confirm the exact workbook is closed, validate the M source, preserve a recoverable baseline, use the approved sync mechanism, and re-extract to compare. Keep workbook features, existing tables, load bindings and unrelated queries intact. A query test requiring Excel evaluation needs its own explicit refresh instruction. Follow [the runner instructions](ResidentialCare-Runner-Agent-Instructions.md) for any runner validation or execution.

For every test, record the source revision, exact target, query evaluated, input versions, expected result, actual result and remaining errors. Validate edited M syntax, headers, named steps, comments, dependencies, final `in` targets and public interfaces. Require an actual query evaluation for a runtime pass; a parser result or saved worksheet value alone cannot supply it. Preserve the last accepted version. If a test fails, investigate or revert that single change through the approved process; do not add another repair to the same test.

## Sequence of changes and tests

### 0 Confirm the test inputs

Reuse the existing evidence and confirm that the four inspected workbooks and their approved sources have not changed. Read-only inspection of these additional exact inputs is needed when their changes begin:

- `2. Calculations/Settings Data.xlsx`: current roster dates and role/shift calendar.
- `1. Input/whiddon_availability_leave_extraction.xlsx`: actual source object, headers and identifier/date types.
- `1. Input/Worker's Report.xlsx`: actual contract source object and required fields.

These additional workbook reads are proposed scope for implementation, not authorization supplied by this plan. Inspect metadata and the necessary records before choosing an import mapping. Preserve the existing client, unit, files and paths.

### 1 Repair the reconciliation membership adapter

**Source target:** `2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m`.

**Workbook for an authorized sync or query test:** `2. Calculations/Capacity-ShiftAvailability.xlsx`.

Adapt the current `Employees_TABLE` interface in a named preparation query after a pure named IMPORT. Preserve the published reconciliation interface. The old internal meanings correspond to:

| Availability meaning | Current reconciliation field |
| --- | --- |
| Worker appears in allocations | `In Roster` |
| Worker appears in availability records | `In Availabilities` |
| Worker has an HR record | `In Records`, kept separate from availability membership |

Carry leave membership separately. The user selected keeping leave-only workers in the model under the existing role, name, contract and availability rules. Preserve employee identifiers, facility keys, names and complete role values. This changes membership compatibility, without adding a leave-only exclusion or changing how available hours are calculated.

**Test 1:** evaluate only the import and membership preparation against the current reconciliation table. Require valid source flags, correct membership meaning and no duplicate employee/facility rows introduced by the adapter. Verify the four named workers and Majal Dias as a control. This test must not depend on the still-broken availability-record or contract imports.

**Current result:** the exact workbook/source gate passed. The source now adapts the current membership fields and retains leave-only membership, with no new production check queries. The isolated native test could not start because the installed Power Query SDK test executable cannot load `System.Text.Encodings.Web, Version=8.0.0.0`. No M runtime pass is claimed. Workbook synchronization and refresh have not been performed; Change 2 has not started.

**Stop for review.** A full Availability refresh is not the acceptance test for this isolated change.

### 2 Repair the availability-record interface

**Target:** the same Availability source, in a separate revision.

Inspect the exact availability/leave source and establish why the expected `Payroll Code` field is unavailable. Map the approved source object and actual fields explicitly in preparation after the named IMPORT. Preserve employee codes as identifiers, record reasons, start/end timestamps and facility keys. Keep the existing availability and leave calculations unchanged.

**Test 2:** evaluate only source navigation and record preparation. Verify identifier coverage, readable timestamps and reasons, row counts and required columns against the source. Distinguish a worker with no source records from a failed import. Require the missing-`Payroll Code` error to be resolved without changing Test 1's accepted membership result.

**Stop for review.** Do not guess a header, select an arbitrary sheet, or suppress an import error with an empty table.

### 3 Verify and repair the Availability contract reader if needed

**Target:** Availability's Worker's Report import only.

Inspect the approved Worker's Report object and evaluate Availability's reader independently. StaffListMaster's saved `Table1` error establishes its own failure; Availability's reader must be tested separately. If Availability already reads the source correctly, record a pass and make no change. Otherwise repair only its object navigation or schema adapter.

**Test 3:** verify employee identifiers, fortnightly contracted hours, duplicates, missing values and permitted zero-hour casual records. Preserve the existing contract and Settings fallback rules. Require one explicit source object and no arbitrary first-table fallback.

**Stop for review.** A corresponding StaffListMaster repair remains a separate change.

### 4 Test Availability as a complete producer

**Code change:** none, unless a separate defect is identified and reviewed.

After Tests 1 to 3 pass, refresh only the exact Availability workbook under the authorized test instruction. Verify `ReconciledWorkers_CHECK`, `ResDayShift_CHECK`, `Availability-StaffList` and `ResDayShift`.

**Test 4:** require successful evaluation and current roster dates, unique employee/facility and worker/date/shift keys, correct roles, and positive availability supported by the current source rules. Trace the four workers by employee identifier and explain any legitimate absence of available shifts. Do not force a positive row or add test-worker filters to production M. Cached July/August outputs cannot pass this test.

**Stop for review.** Keep downstream consumers unchanged until this producer is accepted.

### 5 Repair and test the StaffListMaster contract reader

**Source target:** `2. Calculations/StaffListMaster.xlsx_PowerQuery.m`.

**Workbook for an authorized sync or refresh:** `2. Calculations/StaffListMaster.xlsx`.

Repair only the failed Worker's Report object/schema interface established from the exact source. Preserve reconciliation as the worker-population authority and retain existing Resource assignment, preferred-role and contract-cap rules.

**Test 5:** test the reader first, then refresh this workbook from the accepted Availability result. Require evaluated, passing required `ResourceContract_CHECK` checks and a current `Table_Masterlist`. Reconcile membership and names with accepted Availability, verify the four workers' intended role coverage, and validate caps against current roster dates and Settings.

Track workers by employee identifier and facility when comparing versions. A regenerated numeric Resource index, including the old control value 80, must not be treated as a permanent employee identifier.

**Stop for review.** Do not add the four names manually or force identity checks to pass.

### 6 Test each downstream workbook separately

Use the accepted producer versions and the approved runner prerequisites. Refresh one workbook, inspect its saved results, and stop before the next row.

| Test | Workbook | Acceptance evidence |
| --- | --- | --- |
| 6A | `2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx` | Current availability maps to Settings periods; the 4,190 baseline period failures are remeasured and no unexplained period failures remain. Reconcile the four allocation identities and skipped quantities. Check original availability, allocation evidence and exclusion reasons. |
| 6B | `2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx` | C# derives from accepted A.1 rows; metadata and eligibility are complete; required checks pass and excluded spare does not reappear. |
| 6C | `2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx` | C### traces to accepted inputs; cap, lineage and publication checks pass; legitimate allocation evidence is preserved. |
| 6D | `2. Calculations/Capacity.xlsx` | AIN capacity reconciles to the accepted B output at matching Resource/Period grain. Identify any unrefreshed other-role inputs; do not call the combined BD output fully current on an AIN-only test. |

These are refresh tests, not proposed calculation edits. A failure starts a new, separately reviewed change. Retain existing allocated-calendar breach diagnostics and distinguish them from new spare-induced failures.

## Later runner change

Once the calculation chain is accepted, propose one separate runner change that detects failed required producer checks and stale or incompatible roster dates. Test it with controlled success, failed-check, missing-output and mixed-period cases before a live run. The test must prove that the observed failure cannot receive a successful business-validation result. Keep this separate from import repairs and from any full-run request.

## Completion and next action

Completion requires accepted current producers, explained identity coverage, reconciled AIN outputs and a record of each test. A nonzero capacity total alone is insufficient. Other units and a full multi-role BD run remain separate scopes.

The current stop point is **Change 1 source review and its pending runtime test**. Resolve the test-engine dependency or explicitly authorize the exact workbook sync and isolated query evaluation before proceeding to Change 2.
