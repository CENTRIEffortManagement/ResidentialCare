# OvertimeRosterChecks implementation plan

This records the approved implementation plan and the subsequent requirement to minimise queries and checks while keeping transformations accessible through named Applied Steps. Business requirements are in [ResidentialCare overtime and roster checks](ResidentialCare-Overtime-and-Roster-Checks.md).

## Power Query structure

- Follow `AGENTS.md`, the mandatory Power Query M skill, and its import/output-contract and M-code-style references.
- Minimise the number of new queries. Create separate queries only for external-source landing, the agreed scalar parameters, distinct calculation responsibilities or row grains, and published outputs. Do not add a preparation query for every small transformation.
- Within each query, expose transformations as clear, meaningfully named `let` steps that can be selected in Power Query's Applied Steps. Keep each distinct transformation in its own step; avoid nested transformation chains and a large record-returning model that hides the pipeline.
- Keep `IMPORT`, `INPUT`, `EXTRACT` and `LINK` queries unchanged landing/navigation only. Put transformations and necessary checks downstream.
- Keep checks proportional to the approved risks: required inputs, unique identity/matches, configuration, coverage and output conservation. Reuse each result instead of duplicating checks across consumers. Keep synthetic test assertions in the test harness rather than adding production audit queries.
- Preserve existing query interfaces and calculations except for the approved enrichment and new output. Review row growth, column width, repeated evaluation and buffers before completion.

## Calculation and outputs

Use the copied Worker Reconciliation workbook shell for the independent organisation-level calculation workbook. Author its source at `Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq`; target `CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx`.

Read published rosters and Worker’s Reports for BD, TE and JH-RY, existing facility/role mappings, Settings Data and SharePoint thresholds. Reuse each Settings Data import for ShiftPeriod and PermutationDimensions. Use all supplied employee duty across configured units and roles for rest and duty-free-day checks. Keep the initial BD plotting population, including existing unmatched Master Roster plotting rows.

Calculate four groups: shift net hours; minimum break and short-gap review; fortnight duty-free days and overtime review; and chronological fortnight hours above 76. ADO and the pre-night-shift check are excluded. Keep the AM/PM/NS label as shift context. Count all net rostered hours once in the 76-hour counter, including hours also identified as daily excess.

Use the earliest Date across BD, TE and JH-RY Settings Data PermutationDimensions tables as one common fortnight start date. Select the minimum date rather than relying on row order or the first worked shift. No manual SharePoint date-anchor setting or Published Roster fallback is used. Calendar completeness remains a separate input requirement.

Publish `OvertimeRosterChecks_ShiftResults` at one row per physical published shift, fortnight results at employee/payroll-fortnight grain, and excess segments in net-work-hour positions. Keep a shift crossing a fortnight boundary as one shift with separate period contributions. Preserve daily and fortnight triggers and count their overlapping excess once. Exact clock intervals require meal timing.

Use agreement stream, grade and documented exceptions for eligibility. Return explained Unknown results for missing eligibility, unresolved meal treatment, incomplete duty coverage, missing preceding duty and ambiguous identity. Do not infer a 15% penalty from shift labels. This version produces roster-review evidence, without monetary costing or deciding which shifts payroll must pay after a days-off shortfall.

## Shifts handoff and Tableau

Introduce deterministic `SourceShiftID` values using employer, employee identity, facility, original boundaries and source role. Validate employee aliases. Match Shifts by unit, allocation employee code, role and original start/end; require a unique result and do not match by name.

Enrich BD Shifts before interval expansion. Carry shift identity through StartNames, EndNames, StaffShifts and NAMES+INTERVALSLIST. Pair each start to its own finish boundary and role rather than another row with the same display name. Unpivot only the four plotting timestamps. Do not fill down new nullable check attributes between shifts. Retain unmatched plotting rows with Unknown check results. Reuse one legacy-measure snapshot at local-shift grain for plotting and interval outputs; ambiguous legacy Name-based classifications remain Unknown rather than borrowing another employee's result.

Publish the Tableau-facing `OvertimeRosterChecks` query from Shifts at four plotting points per existing shift. Keep `Name`, `Role`, `DateTime`, `StaffCount` and legacy duration-band `ShiftType`, plus identity, coverage, match and check attributes. Use separate `RosterShiftPeriod` for AM/PM/NS. Retain other legacy StaffDoubleShifts fields internally where existing consumers need them.

Keep ResourceIntervalAllocation in AllocationByShiftAverage; its existing NAMES_INTERVALSLIST import carries the new attributes. Preserve the interval grid and IDs. SourceShiftID identifies a shift and repeats on plotting and interval rows; it is not a one-to-one relationship key between these outputs. Shift-level measures must not be summed over repeated plotting/interval rows.

Dependencies remain one-way: source inputs → independent calculation → Shifts → ResourceIntervalAllocation. The new calculation must not import Shifts or its downstream workbooks. The user owns TableauConnect routing and manual Tableau replacement. Do not edit Tableau charts, relationships, workbooks or Cost.

## Threshold import and activation

Create `docs/OvertimeRosterCheck/OvertimeRosterThreshold.xlsx`, one sheet and one Excel table named OvertimeRosterThreshold, with `Title`, `Value`, `Unit` and `Notes`. Use unique titles, typed numeric values, explicit units, no formulas and no decorative sheets.

Defaults: daily 10 hours; CS Hospitality 8 hours; CS Personal Care 10 hours; minimum break 8 hours; short-gap review 10 hours; minimum fortnight duty-free days 4; fortnight hours 76; fortnight length 14 days. There are eight numeric settings and no date-value column. The manual anchor and pre-night threshold are removed.

The user creates/uploads the SharePoint OvertimeRosterThreshold list. Use the agreed LINK → filter Title → drill into Value scalar parameters. Keep employee classifications/exceptions separate. The derived roster start date does not establish complete duty coverage; do not infer coverage from first and last duty dates.

## Source gates and verification

Extract the exact approved workbook/source pairs before workbook-linked edits and retain hashes/evidence. Product source belongs under Workflows. Copy the named donor as a shell, without automatically synchronising or refreshing it. Synchronisation, save/close and targeted refresh remain distinct operations governed by AGENTS.md; no production-wide refresh automation is included.

Run native Power Query fixtures against actual source definitions with external inputs replaced. Cover thresholds, meals, cross-unit/overnight duty, midnight and fortnight boundaries, complete/incomplete coverage, minimum-date selection in unsorted/repeated PermutationDimensions rows, missing calendar dates, 74+6→4 excess, overlapping triggers, first-shift Unknown, duplicate/ambiguous identities, missing numeric parameters, and unknown eligibility. Verify plotting-point conservation, existing measures and duration bands, interval keys, metadata propagation, removal of the pre-night outputs and absence of circular imports. Report native tests separately from unsynchronised or unrefreshed workbook runtime.
