# OvertimeRosterChecks

The implementation follows [the approved plan and Power Query structure constraints](../../../docs/OvertimeRosterChecks-Implementation-Plan.md). New queries use named Applied Steps; small transformations stay within their query. The BD enrichment adds four queries: import, shift preparation, one reusable legacy-measure snapshot, and the plotting output. It preserves the existing interval grid and legacy duration formulas.

## Exact source/workbook pairs

| Approved source | Workbook target | Current deployment state |
|---|---|---|
| `Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq` | `CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx` | 33 queries installed, saved and closed; exact re-extraction matched. Three named output loads installed. The current live refresh is in progress. |
| `Workflows/ResidentialCare/OvertimeRosterChecks/BD/Shifts.pq` | `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shifts.xlsx` | 48 queries verified after synchronisation, save and close; exact re-extraction matched. The new plotting output load is installed. Live refresh has not started. |

Baseline hashes and extraction evidence are in `outputs/overtime-roster-checks/source-gate/`. The adjacent historical `.xlsx_PowerQuery.m` files were not edited. The fresh AllocationByShiftAverage extraction confirms its existing ResourceIntervalAllocation transformations carry additional input attributes without another source edit.

The calculation copy's inherited query loads were replaced with the approved core definitions and three output loads, and its FilePathUrl was recalculated for the new target. Existing BD Shifts outputs were preserved while adding the plotting output. Installation and formula-only integration-fix receipts are in [the current deployment folder](../../../outputs/overtime-roster-checks/deployment/20261104-202652/); the consolidated state is in [implementation-verification.json](../../../outputs/overtime-roster-checks/implementation-verification.json). The installation saves and source comparisons passed; live output verification remains pending.

## Data handoff

The independent calculation reads the three configured units' published roster sheets, Worker’s Report sheets, Settings Data ShiftPeriod and PermutationDimensions tables, existing SharePoint role/facility mappings, and OvertimeRosterThreshold. It has no dependency on Shifts or any downstream workbook. No attendance/history feed is read.

The fortnight start is the earliest Date across the BD, TE and JH-RY PermutationDimensions tables, selected with a minimum rather than relying on row order. It is a common roster-period anchor across units. The manual SharePoint date anchor and the pre-night-shift check have been removed. AM/PM/NS remains a classification label. Source export coverage is still required independently of the calendar start date.

The core outputs are:

- `OvertimeRosterChecks_ShiftResults`: one physical published shift; identity, measurements, eligibility, coverage and review results.
- `OvertimeRosterChecks_FortnightResults`: one employee/roster-fortnight; retain period-specific totals here.
- `OvertimeRosterChecks_ExcessSegments`: excess positions within net work, preserving applicable triggers and their union. Clock bounds require sufficient meal timing.

In BD Shifts, `RosterShiftsWithChecks` matches unit, allocation employee code, role and original boundaries. It retains unmatched or ambiguous rows with `OvertimeMatchStatus` and an explanation. Ambiguous matches never multiply the plotting rows. `SourceShiftID` and nullable checks are carried from the shift before interval expansion; new attributes are never globally filled down. `LocalShiftKey` identifies a local input row, including unmatched Master Roster rows; it includes an execution-local row number and is not the stable cross-workbook published-duty ID.

Interval starts pair with their own original finish and role, rather than the next boundary sorted by display name. Missing or ambiguous grid finish boundaries fail explicitly. `LegacyShiftMeasures` accepts one distinct legacy result per local shift and one employee code per display name. Where old Name-based calculations cannot distinguish employees, `LegacyShiftStatus` and `LegacyShiftReason` explain the unknown legacy classification. The new export still has four points per local input shift; it does not reproduce multiplication from legacy Name/time joins. Legacy formulas remain available in StaffDoubleShifts, and unambiguous legacy results remain unchanged.

The public `OvertimeRosterChecks` table retains `Name`, `Role`, `DateTime`, `StaffCount` and legacy `ShiftType`, followed by the check attributes. It has four plotting points per existing shift. `RosterShiftPeriod` is separate from the duration-band ShiftType.

`NAMES_INTERVALSLIST` carries the same attributes into the existing `ResourceIntervalAllocation` table in AllocationByShiftAverage. `SourceShiftID` identifies a physical shift; existing interval IDs identify the unit's current grid. SourceShiftID repeats on both plotting and interval rows. Use the unique-grain core helpers for totals, rather than summing repeated shift/FN hours in these expanded outputs. Review flags on an interval row describe its source shift; they do not declare that the entire interval is payable overtime.

## Measurement and column contract

| Fields | Meaning and units |
|---|---|
| `NetRosterHours`, `QualifyingNetHours`, `BreakMinutes` | Source net roster hours, the counter's hours after supported availability-meal treatment, and source meal minutes. Deduct unpaid meals once. |
| `DailyHoursLimit`, `DailyExcessHours`, `DailyHoursFlag`, `DailyStatus` | Configured limit and excess in hours, with a nullable trigger and applicability/data status. |
| `PreviousSourceShiftID`, `PreviousShiftFinish`, `GapHours` | Previous supplied employee duty across units, its finish, and verified elapsed hours. `GapWallHours` is a local-clock diagnostic; clock-change ambiguity leaves elapsed hours Unknown. |
| `MinimumBreakFlag`, `ShortGapReviewFlag` | Separate roster-rest and payroll-review triggers. Overtime exemption does not suppress the roster-rest measurement. |
| `PayrollFortnightStart`, `PayrollFortnightEnd`, `PayrollFortnightCount` | Existing output names retained for compatibility. Boundaries are roster-based, anchored to the earliest Settings Data PermutationDimensions Date; start is inclusive and end exclusive. A shift touching multiple periods has null scalar start/end and before/after counters; its contributions remain in the period helpers. |
| `FortnightHoursBefore`, `FortnightHoursAfter`, `FortnightExcessHours` | Net roster hours in chronological order; excess is the shift's contribution above the 76-hour counter. Daily excess still enters that counter once. |
| `DutyFreeDays`, `DaysShort`, `DaysOffFlag`, `DaysOffOvertimeReviewFlag` | Whole local duty-free dates and days short, with separate roster and payroll-review flags. For a multi-period shift, scalar DutyFreeDays is null and DaysShort is the maximum known shortfall across touched periods. |
| `RosterShiftPeriod` | AM/PM/NS classification label. It does not produce a pre-night-shift rule or flag. |
| `OvertimeHours`, `TriggerSummary`, `ResultReasons` | Union of daily/fortnight excess in hours only when both can be resolved, preserved trigger labels, and explanations for incomplete facts. Short-gap/days-off review flags do not invent payable intervals. |
| `IdentityStatus`, `MatchStatus`, `CoverageStatus`, `EligibilityStatus`, other status fields | Verification/applicability of the individual inputs and results. Nullable flags distinguish Unknown from false. Exemptions and missing payroll facts do not remove supplied duty from the roster calendar. |

The adapter recognises optional roster columns `MealAvailabilityRequired`, `MealStart`, `MealEnd`, `CoverageStart`, `CoverageEnd`, `AgreementStream`, `Grade`, `OvertimeExempt` and `BrokenShiftException`; payroll policy fields can also come from the worker reports. Missing columns remain null. Conflicting supplied roster/worker policy facts remain Unknown rather than selecting a convenient value. `MealAvailabilityRequired` describes the supplied aggregate break; mixed meal treatment needs a richer confirmed input. Clock mapping currently supports one timed unpaid meal. A coverage end supplied as a date is inclusive; a datetime coverage end is exclusive. These contracts do not authorise changing source exports.

## Deployment status and remaining handoff

1. The existing SharePoint `OvertimeRosterThreshold` list has eight numeric settings. The matching local import template is `docs/OvertimeRosterCheck/OvertimeRosterThreshold.xlsx`, with Title, Value, Unit and Notes. No manual date-anchor or pre-night threshold is read. No SharePoint content was changed during deployment.
2. Supply complete roster-export coverage for all configured units. Missing duty dates do not prove coverage, and the Settings calendar does not establish complete employee duty records. Supply authoritative agreement stream, grade, exemptions, broken-shift facts and meal availability/timing where needed. Optional adapter fields are documented in the source; their absence produces Unknown results and does not authorise editing the original exports.
3. Synchronisation of both exact pairs is complete under AGENTS.md. The core has `OvertimeRosterChecks_ShiftResults`, `OvertimeRosterChecks_FortnightResults` and `OvertimeRosterChecks_ExcessSegments` loaded on `ShiftResults`, `FortnightResults` and `ExcessSegments`; BD Shifts has the new `OvertimeRosterChecks` plotting load. Both installation saves were closed and their re-extracted definitions matched the approved sources. See the [core receipt](../../../outputs/overtime-roster-checks/deployment/20261104-202652/Core-sync-receipt.json) and [BD Shifts receipt](../../../outputs/overtime-roster-checks/deployment/20261104-202652/Shifts-sync-receipt.json). The live Facilities connector exposes `Location`, not `Title`; the downstream mapping and its synthetic fixture were aligned, retested and synchronised without changing thresholds.
4. The authorised targeted refresh/save/close sequence is the core calculation, BD Shifts, then BD AllocationByShiftAverage. The current core attempt is refreshing its first loaded output under the existing 30-minute deadline; the two downstream refreshes have not started. Successful live save/close and final post-refresh source comparisons remain pending. Upstream unit inputs, Settings Data and Intervals are read as saved; no unselected upstream workbook is refreshed. Production runner definitions have not been changed to schedule this dependency.
5. The user routes the Shifts `OvertimeRosterChecks` export through TableauConnect and replaces the old plotting table. SourceShiftID is available on the plotting and interval outputs; the old Name/Role relationship remains coarse until the user changes it. No Tableau charts, relationships, report workbooks, Cost or SharePoint content were modified by this implementation.

## Tests

Run `scripts/test-overtime-roster-checks-native.ps1` for core synthetic cases and `scripts/test-overtime-shifts-native.ps1` for baseline/enriched plotting and interval-consumer compatibility. Both harnesses use actual M definitions with explicit synthetic input substitutions and reject external connectors. Native execution requires the installed Microsoft Power Query SDK and a functioning Windows-profile DPAPI context. Test evidence is written beneath `outputs/overtime-roster-checks/`.

Native fixture success does not establish workbook synchronisation, live refresh, source coverage or payroll approval. Those states must be reported separately.

The current Facilities-schema integration fix passed 57 core native assertions. The unchanged BD Shifts source and existing allocation consumer retain their 29 passing assertions, for 86 total and zero native data sources. Live refresh results must be assessed separately from those fixtures.
