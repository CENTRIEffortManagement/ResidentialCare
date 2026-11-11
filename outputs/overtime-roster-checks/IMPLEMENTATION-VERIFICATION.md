# OvertimeRosterChecks implementation verification

Source implementation and synthetic native validation passed, including the 5 November 2026 amendment removing the pre-night check and deriving the fortnight start from Settings Data. Both approved source/workbook pairs have been synchronised, saved, closed and verified by exact re-extraction. The current core live refresh is in progress; downstream refresh and live output verification remain pending.

| Verification | Result |
|---|---|
| Core calculations | 57/57 native assertions passed using the actual source after aligning the Facilities mapping and fixture to the live `Location` column. |
| BD Shifts and existing ResourceIntervalAllocation consumer | 29/29 native assertions passed, including unchanged unambiguous legacy measures, same-name employee collisions, four plotting points, nullable attributes, and adjacent double-shift legs. |
| Native data-source analysis | Zero data sources in both suites. Landed inputs were synthetic; no workbook or SharePoint reads occurred during native evaluation. |
| Threshold workbook | One sheet, one actual Excel table, eight unique numeric settings, four columns and no formulas or blank settings. All retained numeric values were preserved; structural and visual verification passed. |
| Fortnight start and removed rule | Earliest Date across the three configured Settings Data PermutationDimensions tables. Tests cover later first worked shifts, unsorted/repeated dates, empty calendars and removal of pre-night output fields/triggers. |
| Project source check | `validate-pq.ps1` passed against `pq.project.json`; 19 Power Query files checked. |
| Source-of-truth gate | Fresh read-only M extractions and workbook hashes recorded for the named donor, BD Shifts, copied shell and existing allocation consumer. |
| Source/workbook synchronisation | Core: 33 queries and three output loads installed. BD Shifts: 48 queries verified, with the new plotting load installed and existing outputs preserved. Both saved/closed and re-extracted definitions matched. |
| Current live refresh | Core first loaded output is refreshing under the existing 30-minute deadline. BD Shifts and AllocationByShiftAverage refreshes have not started. No successful live save/close or final post-refresh source comparison is claimed yet. |
| SharePoint, Tableau and Cost | No content, routing, report, relationship or costing changes. The existing threshold list was read/verified; no upload was performed. |

The current 86 native assertions, source/template hashes and deployment state are recorded in [implementation-verification.json](implementation-verification.json). The Facilities integration fix passed a fresh 57-case core run; the unchanged BD Shifts source retains its previous 29-case passing run. The preceding implementation's 82-case record is retained in [the previous verification](implementation-verification-before-20261105-amendment.json). Earlier failed receipts retain evidence of corrected fixture conflicts and invalid-duty cases; the current source hashes match passing receipts.

## Resulting source

- `Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq`: 33 queries, including eight numeric settings, the derived roster start date and three unit-parameterised file landings. Settings Data is reused for shift-period labels and calendar dates. Named steps expose preparation, chronology, work segments, period contributions and the three outputs.
- `Workflows/ResidentialCare/OvertimeRosterChecks/BD/Shifts.pq`: four new queries for the unchanged core landing, physical-shift preparation, one reused legacy-classification snapshot and the public plotting output. The snapshot is necessary to avoid Name/time fan-out without repeating its grouping logic or publishing unused legacy columns.
- `docs/OvertimeRosterCheck/OvertimeRosterThreshold.xlsx`: SharePoint import template with Title, Value, Unit and Notes. The obsolete manual anchor, DateValue column and pre-night threshold were removed.

Unknown results remain explicit for missing classifications, meals, preceding duty, incomplete coverage and invalid/conflicting records. The 76-hour counter includes daily-excess hours once. Daily/fortnight segments preserve their union and split around known unpaid meals. Short-rest and days-off payroll flags do not invent payable hours.

The identity tests justified pairing interval starts to their own finishes, rather than the next boundary sorted by Name. The public output derives its four points directly from StaffShifts. Legacy formulas remain in StaffDoubleShifts; ambiguous Name-based legacy classifications expose Unknown rather than borrowing another employee's result.

## Deployment evidence and live verification

The approved core definitions and three named output loads are installed in the calculation copy, replacing inherited donor query loads while retaining its static/path features. BD Shifts contains the approved enrichment and new plotting load. Installation saves and exact formula comparisons passed. See [the deployment folder](deployment/20261104-202652/), [core synchronisation receipt](deployment/20261104-202652/Core-sync-receipt.json), [BD Shifts synchronisation receipt](deployment/20261104-202652/Shifts-sync-receipt.json) and [the handoff](../../Workflows/ResidentialCare/OvertimeRosterChecks/README.md).

An initial live core attempt failed without saving because the Facilities connector exposes `Location` rather than `Title`. The downstream mapping and its native fixture were corrected, all 57 core assertions passed, and the formula-only update was saved, closed and re-extracted with exact parity. The earlier receipt and verified workbook backup remain in the deployment folder. Threshold settings were unchanged.

The resumed core refresh is currently evaluating its first loaded output. BD Shifts and BD AllocationByShiftAverage have not been dispatched. All successful live output/save/close evidence and final post-refresh source comparisons are pending. The selected sequence reads unselected upstream roster, worker, settings and interval workbooks as saved; it does not refresh them. Production-wide runner configuration was not changed.

The existing SharePoint threshold list is available, but export coverage and authoritative classification/meal facts still determine whether individual results can be resolved. The common fortnight start is derived from Settings Data; the existing PayrollFortnightStart/End output names are retained for compatibility and now represent roster-based boundaries. Named synchronisation and targeted refresh remain separate states under AGENTS.md. SharePoint content, TableauConnect routing, Tableau reports/relationships and Cost remain untouched.

Structural performance review narrowed buffers and replaced global per-row work/duty scans with key joins. Plot points, interval endpoints and role expansion deliberately repeat shift attributes; totals must come from the unique-grain helpers. Separate loaded output queries can reevaluate upstream data independently. The current live attempt has not completed, so no completed full-volume refresh duration or memory profile is reported.
