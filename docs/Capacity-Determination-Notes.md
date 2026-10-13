# AIN capacity determination: A.1, A.2 and B

Reviewed: 5 October 2026. This guide explains the freshly extracted BD/AIN Power Query sources. It records proposed improvements for the next task; no calculation code or workbooks were changed during this documentation update.

Proposals updated: 6 October 2026. The [detailed repair list](Capacity-Determination-Repair-Proposals.md) distinguishes original-code problems from recent changes and applies the cluster/break rule before further B reductions.

## The story

The three files turn worker availability into modelled capacity for each date and shift. **A.1 prepares the choices, A.2 applies the first reductions, and B applies further reductions and publishes C###.** Existing roster allocations inform those choices. The chain calculates capacity; it does not edit those allocation records or create a new roster.

The [dependency diagram](Capacity-Determination-Lineage.mmd) shows the source queries and the workbook interfaces they expect. The [Unit1 contract handover](Unit1-AIN-Contracted-Hours-and-Availability.md) explains the September contract-cap rollout and exact-hours over-contract reporting, with its historical scope preserved.

## What comes in

These are the inputs named by BD A.1's imports. Their loaded contents and upstream BD implementations were not inspected in this review.

| Input | What it contributes |
| --- | --- |
| `Capacity-ShiftAvailability.xlsx`, `ResDayShift` table | Eligible worker/date/shift rows. A.1 assigns nominal availability of 1 per source row. |
| `StaffListMaster.xlsx`, `Table_Masterlist` table | Worker-to-Resource mapping, role and the employee's `Effective Shift Cap`. B imports this independently too. |
| `AllocationByShiftAverage.xlsx`, `ResourceShiftAllocation` table | Existing allocations, expressed as shift/FTE measures, used to recognise worked days and compare capacity with allocations. |
| `Demand.xlsx`, `ShiftDemandHCAverageANACC` table | Required staffing by role, date and shift, used to calculate period surplus and reduction limits. |
| `Settings Data.xlsx` | The configured date/shift calendar and allocation start date. A.1's retained legacy `MaxAvailability` extraction is not its active cap source. |

A **Resource** is the master list's role-specific worker identifier. A **Period** is a configured date/shift key. Resource totals cover the whole Settings roster, rather than each calendar month separately.

## A.1 prepares the choices

A.1 maps availability and allocations to Resources and Periods, then recognises work patterns from allocated days. Its cluster and day classifications help prioritise which available days to retain or remove. It also chooses between multiple available shifts on a day.

For each Resource, A.1 compares roster availability with the imported effective cap and calculates a reduction target. It combines that target with demand surplus, day classifications and priorities to prepare candidate removals. It publishes original availability, allocations, demand, the period grid and reduction candidates for the next files.

A day marked as an off-day removal candidate is a priority decision. That marking alone does not guarantee the day will be removed later.

## A.2 produces C#

A.2 reads A.1's saved interfaces and selects reductions within both period and Resource limits. It then produces **C#** for each Resource–Period.

It first sets explicitly excluded day/shift cells to zero, including the configured roster-day exclusions, missing original availability and rejected shift choices. For other cells, it adds the selected signed reduction to original availability. If no reduction was selected, the original availability remains.

## B produces C## and C###

B reads C#, A.1's original availability, allocations, demand and day priorities, and the master list's contract caps. Its first subtraction pass produces **C##**. A second subtraction pass produces **C###**, then attaches the original Resource availability and the Resource limit for reporting.

B blocks final publication when its input-key, contract-coverage or availability-lineage checks fail. Those checks establish that the inputs can be used together. They do not currently reconcile every final Resource total to its cap or recheck the combined allocation/spare pattern for the required days off.

Capacity and Effort sit downstream of this chain. The diagram marks that boundary conceptually; exact BD consumer imports and Excel load bindings were not audited here.

## Reading the outputs

| Value | How to interpret it |
| --- | --- |
| C#, C##, C### | Successive versions of capacity. Use the appropriate stage; do not add the three together. |
| Spare | Retained capacity on an unallocated Resource–Period. It is not a separate roster allocation created by this chain. |
| `ResAvailability` | Original availability aggregated over the roster. It can exceed the cap before reductions. |
| `ResMaxAvail` | A Resource-level contextual limit attached to final rows. Its presence is not proof the final total meets it; do not sum the repeated value across Periods. |
| Existing allocation | A separate input measure. Capacity on an allocated cell can overlap it, so adding all allocations to all C### can double count. |

The imported effective cap may be below 20 for an individual worker. The reviewed distribution code applies roster-wide targets; it does not implement an independent limit of 20 for each calendar month. Actual allocated hours and fractional allocation measures also need to be distinguished from nominal availability counts.

## Improvements for the next task

These are proposals from the source review, not implemented changes or findings reconstructed from refreshed BD data.

1. **Repair the recent allocation-check weakening.** Restore real allocation-name coverage failures so worked days cannot disappear before cluster analysis. Retain and verify the recent chronological and Resource-specific cluster fixes.
2. **Establish legal spare before C#.** Repair the inherited classification expressions and choose start/end extensions jointly in A.1, sharing the five-worked-day allowance and protecting two full off days between groups. A.2 must enforce that eligible set independently of demand surplus or cap reductions.
3. **Complete the preparation and interfaces.** Include eligible unallocated single-shift days, preserve required cluster context and provide the actual day priorities to downstream consumers. Confirm counting definitions and saved interfaces.
4. **Apply cap reductions to valid spare.** Count only jointly accepted reductions against period and Resource allowances. B's further reductions must enforce the effective cap even without demand surplus, using demand to rank feasible removals. Preserve allocation facts and distinguish existing breaches.
5. **Verify publication rather than repair illegal spare late.** Gate C# and C### on the prepared decisions, cap and approved work/rest rules. Final rest checks detect an upstream failure; they do not replace upstream eligibility.
6. **Complete tests omitted from earlier validation.** Exercise the full worker calendar and source interfaces, including competing extensions, partial work, single-shift availability, caps below 20, shortages and unmatched mappings. The detailed repair list records the original/recent categories and remaining policy decisions.

## Evidence and applicability

The current explanation is grounded in the exact BD extracts produced on 5 October 2026: [A.1](<../CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx_PowerQuery.m>), [A.2](<../CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AIN/CapacityDistrib(A.2)-shifts.xlsx_PowerQuery.m>) and [B](<../CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx_PowerQuery.m>).

The comparison found matching executable A.1 code in BD and TE; Unit1 retains an earlier contract-aware A.1, and the user-authorised Unit2 reference retains the older Settings-cap approach. Stored A.2 and B decision code is shared across the compared copies, but several extraction headers identify other or legacy workbooks. In particular, Unit2 A.1 was used as the existing reference at the user's instruction. These comparisons do not establish current source-to-workbook parity for every unit.

This documentation review did not inspect workbook contents, synchronise M code or refresh Excel. Saved output data, actual load bindings and the reported BD totals remain matters for the subsequent issue-solving task.
