# DATExx-Whiddon TE Worker Reconciliation Data Flow

Status: corrective employee-level employment and termination repair implemented in the adjacent Power Query source on 27 September 2026. This revision has not been synchronized or refresh-verified.

## Scope and authority

This flow applies to `CLIENT/DATExx-Whiddon/UNITS/TE`.

`Worker Reconciliation.xlsx/Employees_TABLE` is the sole authority for worker membership. `StaffListMaster.xlsx` must not independently append allocation workers to that population. Allocation data may still be compared with the reconciled list for diagnostics.

In this document, **Records** means the availability/leave extraction. The local unfiltered `Worker's Report.xlsx/Table1` supplies employment, termination, role and contract enrichment. `Capacity-ShiftAvailability.xlsx` uses the same file for its existing employee-level contract resolution.

## Implemented flow

```text
Allocation/Published Roster ─────┐
Availability and leave records ──┼─> Worker Reconciliation
Worker's Report ─────────────────┘      membership + termination gate
                                            │
                                            ▼
                                      Employees_TABLE
                                            │
Worker's Report contract hours ──> Capacity-ShiftAvailability
                                            │
                                            ▼
                                  Availability-StaffList
                                            │
Settings ────────────────────────> StaffListMaster
                                      contract fallback + caps
                                            │
                                            ▼
                                      Table_Masterlist
```

The existing batch order already supports this dependency direction: Workers, then Availability, then Staff.

## Worker Reconciliation

The full outer reconciliation uses employee and facility identity from:

- Allocation: the local published roster used by the allocation workflow.
- Records: `whiddon_availability_leave_extraction.xlsx/Combined Output`.

Employee codes and facility abbreviations are normalized before matching. Records are reduced to one row per employee and facility while retaining `Employee Name`. The published name is the allocation roster name when present, otherwise the records-source employee name. This allows a Records-only worker to satisfy identity validation without pretending that the worker appeared in Allocation. The `HF` Humanforce Admin system account is excluded because it is not a worker.

Allocation and Records source-presence markers are set before the full outer join. Blank roster vacancies and unassigned shifts are excluded before that join because they have no worker identity. Identified agency workers remain in scope even when their agency roster code does not exist in the employee report.

Each retained worker carries:

- `In Allocation`
- `In Records`
- `Worker Membership`, with exactly one of `Allocation only`, `Records only`, or `Both`

The legacy `Availabilities - Not in Roster` diagnostic is intentionally worker-level: it compares employee IDs without facility and returns Records IDs that occur nowhere in Allocation. The separate reconciliation flow remains employee-and-facility grain. `AvailabilitiesSTATS` uses a coalesced facility from either side of the full outer join; it does not depend on the source `Site Name` column being retained.

Roster role parsing removes only a terminal `AM`, `PM` or `NS` suffix. Non-shift roles such as Clinical Care Coordinator are no longer truncated. Where a worker has multiple allocation roles, the role aligned with the employee report position is preferred; remaining ties are resolved deterministically. The published grain is one preferred worker row per employee and facility, and any unresolved multi-row key remains blocking in `DuplicateWorkers_DIAGNOSTICS`.

### Termination rule

The unfiltered worker report is normalized and reduced to one deterministic row per employee before it is joined to membership. An employment row eligible at roster start is preferred; when none is eligible, the most recently terminated row is retained for termination testing. Employment is matched by employee ID, while the published facility always remains the Allocation/Records facility. This is necessary because valid availability workers may offer availability at a site other than their home employment facility.

`Worker's Report[Termination Date]` is compared with the first date in the published roster:

- termination before roster start: exclude from `Employees_TABLE`;
- no termination date, or termination on/after roster start: retain in the minimal implementation.

Excluded rows remain visible in `TerminatedWorkers_DIAGNOSTICS`. A termination during the roster is not yet applied at individual shift-date grain.

Any membership candidate without an employee-level worker-report match remains visible in `UnmatchedCurrentWorkers_DIAGNOSTICS`. A Records-only missing match blocks publication because the worker report is needed for its role and employment status. An Allocation worker remains publishable because Allocation supplies its identity and role; a missing or zero contract then uses the downstream Settings fallback. Repeated eligible worker-report rows with identical material employment facts are collapsed for ambiguity testing. Conflicting eligible rows remain visible in `AmbiguousCurrentWorkers_DIAGNOSTICS` and block publication.

The row-level review interfaces are:

- `DuplicateWorkers_DIAGNOSTICS`
- `IncompleteWorkerIdentity_DIAGNOSTICS`
- `InvalidWorkerMembership_DIAGNOSTICS`
- `UnmatchedCurrentWorkers_DIAGNOSTICS`
- `AmbiguousCurrentWorkers_DIAGNOSTICS`
- `TerminatedWorkers_DIAGNOSTICS`
- `LeakedTerminatedWorkers_DIAGNOSTICS`
- `WorkerReconciliation_ISSUES`, which consolidates all blocking issues

`WorkerReconciliation_CHECK` summarizes those diagnostics and blocks publication for duplicate worker keys, incomplete identity, invalid membership, missing or ambiguous Current Workers matches, or leakage of an excluded terminated worker. The termination-leak test is independent of the earlier exclusion step.

## Capacity-ShiftAvailability

`Capacity-ShiftAvailability.xlsx` continues to own its existing availability calculations and contracted-hours import. It does not rebuild worker membership.

The following reconciliation fields pass unchanged through `ReconciledWorkers_Eligible` and `Availability-StaffList`:

- `In Allocation`
- `In Records`
- `Worker Membership`

Workers remain identifiable even when their records produce no available shifts.

## StaffListMaster and contracts

`Availability-StaffList` is the only worker-population input to `Table_Masterlist`. The allocation staff import remains diagnostic-only and cannot publish an extra Resource.

For each reconciled worker:

```text
Fallback Contracted FN Hours =
    MaxAvailability × ShiftDuration ÷ RosterFortnights

Contracted Roster Hours = Contracted FN Hours × RosterFortnights
Contracted Shift Equivalent = Contracted Roster Hours ÷ ShiftDuration
Contracted Shifts = floor(Contracted Shift Equivalent)
Effective Shift Cap = min(Contracted Shifts, MaxAvailability)
```

A positive worker-record contract is retained. A missing or zero value is replaced with the Settings-derived fallback. Negative, non-finite or otherwise invalid values remain blocking failures.

`Contract Source` identifies `Worker record` or `Settings maximum fallback`. `ResourceContract_CHECK` requires positive contract hours and a usable effective cap for every published Resource. Fallback use is visible as a warning rather than hidden.

## Compatibility boundaries

The minimal implementation preserves the existing `Source`, `Worker Record Status`, `Worker Contract Issue`, `Effective Shift Cap`, and `Limit Basis` columns so current A.1, B and Effort consumers do not require an immediate redesign. The new membership and contract-source columns are additive.

## Source-level validation evidence

Read-only validation on 27 September 2026 used the saved Worker Reconciliation outputs plus the exact current roster, records and root `1. Input/Worker's Report.xlsx` sources. The repaired transformation model produced 379 unique employee/facility workers: 72 Allocation only, 144 Records only and 163 Both. All six blocking reconciliation checks evaluated to zero failures. Three Records-only workers terminated before the 20 July 2026 roster start were excluded and retained in the termination diagnostic: employee IDs `22670`, `16323` and `23184`.

The root and `Allocation` copies of `Worker's Report.xlsx` differ at source row 3713. The root copy, which is also used by downstream contract queries, repeats employee `26068`; the Allocation copy contains `26532` in that row. Repeated eligible rows with identical material employment facts are treated as one for ambiguity testing. Employee `26532` remains a valid Allocation-only worker and receives the Settings contract fallback because the root report has no matching record. This source discrepancy should be corrected upstream, but it no longer prevents the authoritative membership list from publishing.

## Paths and deployment gate

The edited Worker Reconciliation source uses the standard `FilePathUrl` and `CentriSyncPaths` resolver. The roster, availability and unfiltered worker-report imports are all derived from the resolved unit root; no user-specific external path is retained.

Before synchronization, confirm that `Worker Reconciliation.xlsx` contains one valid `FilePathUrl` input identifying that workbook. Synchronization, refresh, save and close are separate opt-in operations. After any approved synchronization, re-extract each workbook definition and compare it with its approved adjacent M source before refreshing.
