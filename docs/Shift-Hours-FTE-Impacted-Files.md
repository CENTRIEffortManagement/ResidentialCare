# Files affected by shift hours, meals, headcount and FTE

Current Unit1 standard-FTE implementation: [settings source, formulas and validation](2026-09-16-Whiddon-Unit1-Settings-Based-FTE-Conversions.md). Source only; workbook synchronization and refresh remain pending.

Updated: 16 September 2026.

## Scope and evidence

The immediate discussion concerned `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m`. The complete topic also touches roster extraction, interval allocation, capacity, organisation calculations and reporting.

This register uses saved M source, existing reconciliation evidence and Tableau `.twb` text. Workbook names identify the associated workflow; where a source exists, the inspected code is the corresponding `.xlsx_PowerQuery.m` file. Embedded queries, worksheet formulas, loaded tables and refresh currency have not been verified. No workbook was inspected or changed to prepare this register.

An affected file may supply inputs, calculate a measure, or consume changed results. Inclusion does not mean every file needs a code edit. Existing extraction/source-of-truth and opt-in synchronization rules still apply.

Path prefixes used below are repo-relative:

- **U1** = `CLIENT/DATExx-Whiddon/UNITS/Unit1`.
- **ORG** = `CLIENT/DATExx-Whiddon`.
- **ROLE** = each of `RN`, `AIN`, and `AINC4`; each capacity-stage row therefore identifies three separate files.

## Input and unit calculation files

| File under U1 | Connection to the topic | Review or validation needed |
| --- | --- | --- |
| `1. Input/Master Roster.xlsx` | Referenced by Manual Read's `IMPORT Master`; supplies roster hours, net length, start/end times and break fields. | Confirm which source field represents net hours. The exact imported path was absent on 16 September; `1. Input/Allocation/Master Roster.xlsx` exists but is not an approved substitute. |
| `1. Input/Demand-MasterRoster Manual Read.xlsx` | Source updated 16 September: selects `Shift Net Length`, aliases it as `Roster Hours`, and uses net hours for historical sums and productive-hour percentages. Assigns each full row by start time; distributes care targets over 14 days and publishes configured standard-FTE equivalents. Headcount is diagnostic. | Net-field change is source-only, not synced/refreshed. Splitting crossing rows and reviewing exact-boundary/overnight attribution remain separate work. Reconcile hours and percentage denominators after an authorised refresh. |
| `1. Input/Allocation/Published Roster.xlsx` | Exact roster source referenced by Allocation Extraction. | Establish the meaning of `Shift Net Length`, breaks and timestamps. Preserve source-row identity. |
| `1. Input/1-AllocationExtracted.xlsx` | Selects `Shift Net Length` as `Hours`; carries start/end and break fields; also produces staff and role information. | Preserve source net hours through downstream stages. Its check uses an approximate `0.9366` factor that is separate from the principal output. |
| `1. Input/2-DemandExtract.xlsx` | Reads Manual Read's allocation/profile/check publications; converts source FTE to hours using validated Settings standard duration and hours to shift FTE using actual duration; repeats the fortnight over 28 days. | Preserve productive and roster-hour totals, role mapping, timing, zero cells and output interfaces if the upstream units change. |
| `1. Input/whiddon_availability_leave_extraction.xlsx` | Direct availability/leave source referenced by Capacity-ShiftAvailability. | Availability windows and leave must remain distinct from observed roster net hours. Source contents have not been inspected. |
| `1. Input/Worker Reconciliation.xlsx` | Maps availability-source workers to roster identities/roles in Capacity-ShiftAvailability. | Validate worker/role identity so multiple records or roles do not become duplicate capacity or headcount. |
| `1. Input/AvailabilitiesExtracted.xlsx` | Removed 20 September 2026 after a repository-wide reference search found no workbook sourcing it. Capacity-ShiftAvailability reads the raw availability/leave source directly. | External consumers outside this repository were not checked. |
| `1. Input/StaffList Availability.xlsx` | Retained availability input workbook from an earlier workflow. Its saved M section contains no shared query definitions. | Worksheet/manual dependencies are unverified; do not infer its current calculation role from its filename. |
| `2. Calculations/Settings Data.xlsx` | Supplies shift periods, boundaries, standard duration and meal settings; also derives roster hours with a separate half-hour rule. | Define the facility/role/shift timing authority and distinguish gross shift span, standard FTE hours and meal policy. |
| `2. Calculations/Intervals.xlsx` | Establishes temporal boundaries and roster intervals; contains a meal-derived duration with a days/hours inconsistency. | Split time consistently, including overnight boundaries; preserve net-hour reconciliation and resolve the duration-unit inconsistency. |
| `2. Calculations/Shifts.xlsx` | Builds worker intervals, combines some roster spans, recalculates meal-adjusted duration and assigns shift types. | Retain source-row net hours; review long, short, adjacent, overlapping and double shifts; avoid recalculating meals on already-net hours. |
| `2. Calculations/ShiftsRFBI.xlsx` | Additional shift-processing variant with similar meal/duration rules. | Review the same rules if used. Its presence does not establish active workbook consumption. |
| `2. Calculations/AllocationByShiftAverage.xlsx` | Produces worker interval effort and role/resource shift allocation; uses effective/clock ratio, actual shift duration and standard FTE duration for different outputs. | Calculate interval contributions from source net hours and reconcile to role shift totals; keep actual-shift FTE and standard resource FTE distinct. |
| `2. Calculations/Allocation.xlsx` | Publishes role allocation and resource allocation for Effort and other consumers. | Preserve both output contracts and confirm each measure's denominator; review its full-shift filter when interpreting partial workers. |
| `2. Calculations/DemandIntervals.xlsx` | Repeats shift demand across intervals; calculates a meal ratio that the principal demand measure bypasses. | Keep demand's interval and shift units consistent; reconcile diagnostic factors without activating an extra meal deduction. |
| `2. Calculations/Demand.xlsx` | Integrates interval demand and publishes average role/shift attendance to Effort and capacity models. | Confirm weighted-average round trip; review integer-typed diagnostic FTE and retained check-field inconsistencies. |
| `2. Calculations/Capacity-ShiftAvailability.xlsx` | Converts availability windows, leave and roster evidence into worker/shift eligibility and effective-hour allowances. | Keep nominal availability policy separate from actual roster net hours; review partial windows and the configured full-shift allowance. |
| `2. Calculations/Availability.xlsx` | Reconciles available workers with allocated workers. | Validate worker counts and identities after any segmentation or role changes. |
| `2. Calculations/StaffListMaster.xlsx` | Builds the worker/role master list from allocation and availability inputs. | Confirm that split rows and multiple roles preserve worker identity without duplicate or lost workers. |
| `2. Calculations/ROLE/CapacityDistrib(A.1)-shifts.xlsx` | Reads Demand, AllocationByShiftAverage, staff and availability; sets nominal worker/shift availability to 1. | Review headcount versus fractional availability explicitly; preserve common units when comparing demand, allocation and capacity. |
| `2. Calculations/ROLE/CapacityDistrib(A.2)-shifts.xlsx` | Caps and prioritises capacity using A.1 outputs. | Revalidate caps, priorities and totals if the capacity/FTE basis changes. |
| `2. Calculations/ROLE/CapacityDistrib(B)-shifts.xlsx` | Redistributes capacity using A.1/A.2 results. | Revalidate worker and shift totals, constraints and redistributed profiles. |
| `2. Calculations/Capacity.xlsx` | Combines role-stage capacity and availability publications. | Preserve aggregation units and reconcile original, capped and redistributed totals. |
| `2. Calculations/Effort.xlsx` | Combines Demand, Allocation, Capacity and resource allocation into comparison outputs. | Verify comparable units by role/day/shift; distinguish standard resource FTE from actual-shift FTE. |
| `2. Calculations/Shift-StaffDistribution.xlsx` | Reads Shifts; produces shift-type staff distributions and effort. Its effort query deduplicates by worker/date before summing effective duration. | Review how multiple rows/shifts per worker/day are counted and retained when segmentation changes. |
| `2. Calculations/MutliRoleCheck.xlsx` | Reads multi-role availability, shift intervals and settings. Filename spelling is retained. | Revalidate worker/role coverage and multiple-role diagnostics. |

## Organisation calculation files

| File under ORG | Connection to the topic | Review or validation needed |
| --- | --- | --- |
| `2. Calculations/E-O-I/Effort-All.xlsx` | Combines unit Effort and publishes organisation comparisons and availability matrices; applies `7.6 / 8` in capacity/resource-availability branches. | Establish each input's basis before changing the factor; avoid an additional reduction of already-net actual-shift FTE. |
| `2. Calculations/E-O-I/EffortOutcomes.xlsx` | Reads Effort-All and derives effort/outcome comparisons. | Revalidate comparisons when demand, allocation or capacity changes. |
| `2. Calculations/E-O-I/Inefficiencies.xlsx` | Derives shortfall, excess and other efficiency measures from EffortOutcomes. | Revalidate derived categories and aggregated totals. |
| `2. Calculations/Tableau Connection.xlsx` | Publishes Inefficiencies data to reporting. | Preserve reporting fields and reconcile published values. |
| `2. Calculations/Cost/Cost..xlsx` | Reads inefficiency measures and now uses Unit1 Settings standard-FTE hours for costs (source update 16 September). Double dot is part of the filename. | Confirm the input FTE denominator before converting to hours/cost. The Settings path now points to Unit1; see [Cost implementation notes](2026-09-16-Whiddon-Cost-Settings-Based-FTE-Conversion.md). |
| `2. Calculations/Change/AllocationChange.xlsx` | Converts change hours using imported standard shift duration; Effort-All references the change branch. | Reconcile hour/FTE conversions and verify current activation/source paths. |
| `2. Calculations/Change/SS/ResidualAllocation.xlsx` | Converts demand/allocation measures to residual hours using imported shift duration. | Verify that the multiplier reverses the actual input denominator; retained external paths make current operation unverified. |
| `2. Calculations/Change/SS/ResidualAllocation-ALL.xlsx` | Related residual-hours branch with role-specific outputs. | Apply the same unit and source-path review. |
| `2. Calculations/EOW/EffortOutcomeLogXY.xlsx` | Transforms effort/outcome results from EffortOutcomes. | Revalidate transformed results after upstream changes. |
| `2. Calculations/EOW/2DRead.xlsx` | Reads the transformed EOW results and grid thresholds. | Revalidate classifications and report outputs; thresholds are a separate input, not a roster-hours calculation. |

## Reporting files

These are downstream validation targets. Tableau text establishes referenced fields/connections, not current extract freshness. Some connections still point to older locations or packaged data.

| File/location | Connection to the topic |
| --- | --- |
| U1 `3. Report/ShiftProfile.twb` | Worker interval contribution, shift profile and effort comparisons from AllocationByShiftAverage, Shifts, Effort and Shift-StaffDistribution. Includes an older connection for the last input. |
| U1 `3. Report/ShiftProfile-LTPaths.twb` | Related profile report with separate connection paths; validate it independently if used. |
| ORG `3. Report/ShiftProfilex1F-LT.twb` | Related profile variant referencing packaged/older data paths. |
| U1 and ORG `3. Report/Availability Redistribution.twb` | Original and redistributed availability/capacity profiles, including `RoleAvailabilityDevelopedMATRIX` outputs. |
| U1 and ORG `3. Report/Availability RedistributionY.twb` | Related redistribution report variant. |
| U1 and ORG `3. Report/Effort and Inefficiencies.twb` | Demand, allocation, capacity and efficiency comparisons. |
| U1 and ORG `3. Report/Cost-LT_v2025.3.twb` | Downstream cost reporting; verify consistency with the cost conversion. |
| U1 and ORG `3. Report/EOW v07 - AG1-1DS.twb` | Downstream effort/outcome reporting. |

Packaged `.twbx` reports, `.hyper` extracts and bundled workbook copies may also contain saved results. Their internals were not inspected; they are not source-code edit targets inferred from this register.

## Other units, clients and retained variants

- The earlier static reconciliation also searched `CLIENT/DATExx-Whiddon/UNITS/Unit2` and `CLIENT/DATExx/UNITS/Unit1`, plus their organisation/reporting branches. Corresponding files can contain the same topic but different formulas. Do not assume parity or propagate edits automatically.
- Retained variants include `1-AllocationExtracted.Christ.xlsx`, DATExx `2-DemandExtracted-Master-.xlsx`, and the Whiddon Unit2 saved source `2-DemandExtracted-Master-Christ.xlsx_PowerQuery.m`. These include older duration/check logic; exact workbook pairing and active use must be established separately.
- Backups, `ReadX`/copy files, migration exports and diagnostic sources are evidence or variants, not automatically approved replacements for the named current source.
- Shared path helpers, runner files, outcome/threshold inputs and geographic mappings are supporting dependencies. This review has not established that their formulas need changes for the hours/FTE decision.

## Related notes and validation assets

- [Hard-coded 7.6-hour and 456-minute audit](../outputs/hardwired-fte-audit/README.md), including all eight business M-source files and separate test/backup occurrences.
- [Discussion and measure definitions](Demand-Duration-and-Meal-Time-Discussion.md).
- [Static reconciliation register](../outputs/shift-duration-reconciliation/README.md) and [file/query evidence index](../outputs/shift-duration-reconciliation/index.md).
- [File/table connection map](ResidentialCare-File-Table-Connections.mmd) and [shift-duration diagram](Shift-Duration-Chart.mmd).
- [Manual Read handoff](../CLIENT/DATExx-Whiddon/UNITS/Unit1/1.%20Input/MinuteWorker-FTE-Handoff-2026-08-30.md).
- [Demand Extraction contract](Demand-Extraction-Distributed-FTE-Plan.md), [profile chart instructions](MinuteWorker-FTE-Profile-Chart-Instructions.md), and [availability rules](DATExx-Whiddon-Capacity-Shift-Availability-Rules.md).
- `scripts/test-minuteworker-period-allocation.mjs` and `scripts/test-demand-extraction-distributed-fte.py` contain source/arithmetic fixtures to review when the measure contract changes. They were not run for this documentation update and do not execute Excel Power Query.

The principal implementation review begins with source net-hour meaning, common shift boundaries and row segmentation. Follow the effects through both Demand and Allocation, then capacity and reporting. Reconcile each original row's net hours and each role/day/shift total before accepting downstream comparisons.
