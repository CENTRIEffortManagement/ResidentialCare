# ResidentialCare file staging plan

## Purpose

This plan groups the existing Unit and organisation (`E-O-I`) workbook sequences into continuous, restartable runs. It does not change workbook logic or the authoritative order in the runner manifests.

The staging boundaries come from:

- `CLIENT/DATExx-Whiddon/runner/ResidentialCare-UnitWorkbookSequence.psd1`;
- `CLIENT/DATExx-Whiddon/runner/ResidentialCare-OrgWorkbookSequence.psd1`; and
- `docs/ResidentialCare-File-Table-Connections.mmd`.

An independent run means that the batch has one entry gate, runs its files continuously in the listed order, and produces a completed handoff for the next batch. It does **not** mean that a downstream batch can ignore its upstream prerequisites.

## Staging rules

1. Preserve the existing workbook order inside every batch.
2. Complete and verify a batch before releasing its downstream batch.
3. Treat each `UnitN` as a separate run lane. A unit must not read another unit's partly refreshed outputs.
4. Allow demand, allocation and capacity-input branches to progress independently until they reach a reconciliation gate.
5. Role lanes may run independently after R2 is complete. Unit consolidation must wait for every enabled role lane.
6. Organisation batches start only after R4 has completed for every selected unit.
7. Each batch must receive the required support files shown below, even when those files are not refreshed in that batch.
8. A failed batch is restarted from its first file unless a narrower resume point has been explicitly validated against the runner manifest.
9. This document is a staging proposal only. It does not authorise refresh, workbook edits, or M-to-Excel synchronisation.

## Unit staging template

Apply this same template separately to every selected `UnitN` folder.

| Batch | Existing unit sequence | Continuous run | Entry gate | Release condition |
|---|---:|---|---|---|
| **D1 — Demand transformation** | Pre-sequence + 02 | `Demand-MasterRoster Manual Read` → `2-DemandExtract` | New demand data available | Demand extract accepted at R1 |
| **A1 — Allocation transformation** | 01 | `1-AllocationExtracted` | New allocation/roster data available | Allocation extract accepted at R1 |
| **S1 — Shared controls** | 03 | `Settings Data` | Settings/control data available | Settings version accepted at R1 |
| **C1 — Capacity inputs** | Source + 07–08 | approved employee/availability transformation → `Capacity-ShiftAvailability` → `StaffListMaster` | New capacity inputs available; allocation and settings ready before 07–08 | Capacity inputs accepted at R2 without waiting for demand |
| **D2 — Demand preparation** | 04–06 | `Intervals` → `DemandIntervals` → `Demand` | R1 complete | Demand publication accepted at R2 |
| **A2 — Role-allocation basis** | 09–10 | `Shifts` → `AllocationByShiftAverage` | R1 complete | Allocation basis accepted at R2 |
| **A3 — Allocation output continuation** | 11–13 | `Allocation` → `Shift-StaffDistribution` → `MutliRoleCheck` | A2 complete | Allocation publications ready for unit consolidation; this may continue while C2 role lanes run |
| **C2.1–C2.N — Role capacity lanes** | Dynamic role range | For each enabled role folder: `CapacityDistrib(A.1)` → `CapacityDistrib(A.2)` → `CapacityDistrib(B)` | R2 complete | Every enabled role folder's C-tables and checks accepted at R3 |
| **U4 — Unit consolidation** | 23–24 | `Capacity` → `Effort` | R3 complete | Unit demand, allocation and capacity reconcile at R4 |

Demand, allocation and capacity-input transformation branches can progress independently to their gates. The C2 role lanes are also logically independent after R2. Within each C2 lane, A.1, A.2 and B remain serial. The production runner expands C2.1 through C2.N from the enabled roles in the saved client role profile and executes them in that approved order.

`Role 1` and `Role N` are diagram labels, not literal folder names. For the current saved profile they resolve to the enabled role folders in order. A separate role check may compare actual candidate folders with that profile, but a normal production run must not silently add newly discovered folders.

## Organisation and E-O-I staging

The existing organisation order is preserved. The batches are separated at durable dependency handoffs.

| Batch | Existing organisation sequence | Continuous run | Entry gate | Release condition |
|---|---:|---|---|---|
| **O1 — Cross-unit assembly** | 01–02 | `E-O-I/StafMasterList-All` → `E-O-I/Effort-All` | U4 complete for every selected unit; `AllocationChange` and Unit1 `Settings Data` available | Combined staff and effort publications refresh successfully |
| **O2 — Outcomes and inefficiency** | 03–04 | `E-O-I/EffortOutcomes` → `E-O-I/Inefficiencies` | O1 complete | Outcome and inefficiency publications refresh successfully |
| **O3 — Cost branch** | 05 | `Cost/Cost.` | O2 complete; Unit1 `Settings Data` available | Cost and saving publications refresh successfully |
| **O4 — Outcome history and 2D read** | 06–07 | `EOW/EffortOutcomeLogXY` → `EOW/2DRead` | O2 complete; `Grid Thresholds` available | Outcome-log and 2D publications refresh successfully |
| **O5 — Reporting handoff** | 08 | `Tableau Connection` | O2 complete; O3 and O4 complete when a full end-to-end release is required | Tableau-facing publication refreshes successfully |

O3 and O4 are separate downstream branches after O2. They may run independently if the runner and Excel process-isolation controls support it. For the current production sequence, keep the manifest order: O3, then O4, then O5.

## Release sequence

For one unit, the proposed release path is:

`Demand + Allocation + Settings → R1 → Demand and role-allocation preparation → R2 → role-capacity lanes → R3 → Capacity and Effort → R4`

The capacity-input branch does not wait at R1. Once its own source, allocation extract and settings are ready, it can progress to R2 independently. After `AllocationByShiftAverage.xlsx` reaches R2, `Allocation.xlsx` and its downstream allocation outputs can also continue while the role-capacity lanes run.

For multiple units, repeat the complete unit flow within each unit lane. Organisation O1 waits at R5 until every selected unit has released its reconciled Effort output.

## Event-driven progression model

The target operating model is to watch each input family and progress only the affected branch as soon as new data is ready:

- **Demand event:** refresh `Demand-MasterRoster Manual Read.xlsx`, then `2-DemandExtract.xlsx`.
- **Allocation event:** refresh `1-AllocationExtracted.xlsx`.
- **Capacity-input event:** refresh the approved availability and employee transformations, then `Capacity-ShiftAvailability.xlsx` and `StaffListMaster.xlsx` when their allocation and settings inputs are ready.
- **Settings event:** refresh `Settings Data.xlsx` and invalidate downstream readiness for every branch that consumes it.

Each event should carry a data watermark or run identifier. A branch may advance immediately to its next dependency gate, but a gate releases downstream work only when all required branches have accepted outputs for the same unit and compatible watermark. This permits maximum safe progress without combining new demand with stale allocation or capacity data accidentally.

The proposed reconciliation points are:

| Gate | Reconciles | Downstream release |
|---|---|---|
| **R1 — Demand/allocation alignment** | Demand transformation, allocation transformation and settings/control version | Demand preparation and role-allocation preparation |
| **R2 — Role-input reconciliation** | Demand output, allocation-by-shift output, staff/capacity inputs and settings | Independent role-capacity batches |
| **R3 — Capacity reconciliation** | Every enabled role lane and its role totals/checks | Unit `Capacity.xlsx` |
| **R4 — Unit release** | Final demand, allocation and capacity publications in `Effort.xlsx` | Unit ready for organisation processing |
| **R5 — All-unit barrier** | R4 release from every selected unit for the same run/data watermark | Organisation O1 |
| **R6 — Organisation reconciliation** | Cost and EOW outcome-history branches | Tableau/reporting release |

The watcher and trigger layer is a future implementation. It must not be enabled until the workbook workflow, source-of-truth process, watermark rules, reconciliation checks and failure/retry behaviour are approved.

## Batch control record

Record the following for every batch run:

- batch ID and unit name, where applicable;
- manifest and inclusive sequence range;
- prerequisite batches and support-file checks;
- validation result;
- start and finish timestamps;
- completed workbook count;
- failed or stopped workbook;
- log and status paths; and
- whether any workbook was refreshed, saved, or left open.

## Implementation boundary

The existing runner manifests remain authoritative. Before making these batches executable, add named batch selections or dedicated batch manifests that reference the same workbook entries, then validate every selection without opening Excel. Do not duplicate workbook lists into unrelated scripts, because that would create a second source of truth.

See `docs/ResidentialCare-File-Staging-Sequence.mmd` for the batch-level Mermaid map.
