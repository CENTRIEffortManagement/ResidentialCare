# Demand Extract linked roles

Implemented in `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/2-DemandExtract.xlsx_PowerQuery.m`.

The user explicitly designated this `.m` as authoritative and approved implementation. This work edits that exact source directly. No workbook content was inspected, extracted, synchronized, refreshed, saved or closed. Existing user changes were preserved. `CLIENT/DATExx/`, Unit2 and the Allocation/MasterRoster sources were not changed by this implementation.

The requested related-file commit also includes existing user changes in the Unit1 Settings and MasterRoster `.m` sources. Settings now derives distinct roles from `IMPORT AllocationExtracted` instead of its local Roles table, and MasterRoster preserves DC-role spelling. Those changes were reviewed as text and preserved without further source edits. Workbook binaries and unrelated changes are excluded from this source commit. The existing fixed import path in Settings was not modified or validated against a live workbook.

## Role rules

The reference is `AllocationExtraction` in `1-AllocationExtracted.xlsx_PowerQuery.m`: map roster role to `LINK Roles[Role Group]`, then join the group to `LINK RoleAnalysis[Roles]` and retain `Effort Management Analysis = true`.

Demand's inputs are already aggregated by DC role. The adapted sequence is:

1. Prepare the DC-role definitions from `LINK Roles`, limited to the upstream direct-care categories RN/OTHER.
2. Join every `Role Group` to `LINK RoleAnalysis[Roles]` before checking shared DC roles. Keep the existing source site and list IDs. Missing analysis rows remain errors.
3. A DC role with any enabled group must have one consistent group/category/percentage/inclusion assignment. When all its groups are disabled, collapse them into one excluded lookup row. Preserve common attributes and leave the output group null if its labels differ; do not select an arbitrary label.
4. Join source DC roles to this validated lookup and include only logical `true` effort flags. Logical `false` and null are excluded. Leave-balance flags have no effect. DC-role keys are trimmed and case-insensitive; group labels retain configured case after trimming and must match settings labels.
5. Validate paired allocation/profile data at the original DC-role/day/shift grain, then sum into output role groups.

Repeated roster labels with identical DC-role attributes do not multiply demand. Conflicting assignments for an enabled DC role fail, including any mixture of included and excluded groups: the combined source hours cannot be split after upstream aggregation. Entirely excluded groups may share a DC role without blocking demand. Their source hours are counted once in `Distributed FTE Exclusions`. Missing mappings, missing analysis rows and duplicate analysis keys still fail.

This ordering fixes the reported TA/WLC/WLO conflict. The user confirmed that all three groups are excluded from effort management. The native fixture reproduced the original failure before the fix, then confirmed unchanged retained demand and exact excluded-hour totals after the fix. The live list rows were not directly inspected or modified.

The hard-coded RN/AIN/AINC4 list and the forced EN-to-AINC4 conversion are removed. `Demand Role Code` remains available as a lookup function. `Demand Roles` is now derived from the enabled linked definitions. Every enabled direct-care group must have a complete source pattern and matching settings calendar and timing rows; missing coverage is an error, not zero demand.

## Calculation and interface

- `Distributed FTE Source Cells` preserves original roles and validates 42 fortnight day/shift cells per contributing DC role, including proven sparse zeros.
- `Distributed FTE Prepare` groups those cells by facility, output role, fortnight day, shift and source week. `SourceRoles` lists contributors.
- Roster hours remain `SourceFTE * ShiftDuration`. Productive hours are calculated separately for each DC role before summation. Direct-care percentages are not reapplied to roster demand or averaged across roles.
- `Demand Reconciliation Targets` sums validated source-role cells for the actual dates/shifts in each planning fortnight. `DemandExtraction_RECONCILIATION` compares output group totals with those independent targets, including a final partial fortnight.
- `DemandExtraction_CHECK` exposes mapping, inclusion, source-cell and group-coverage failures. `ShiftUnitDemandHRS` blocks publication on error.
- The twelve published columns, Settings period IDs, overnight timing and fractional attendance calculation are preserved. Output row counts follow the enabled groups and the Settings planning horizon.

Keep added staging queries connection-only if these changes are separately synchronized later.

## Settings-driven planning horizon

The fixed 28-day/84-period/two-cycle assumptions were already present before the linked-role changes (introduced in commit `fd644a3`). They represented two repetitions of the 14-day source pattern; they did not come from a Settings day-count parameter.

The authoritative Settings `.m` defines the horizon through `DateFrom`, `Min Date`, `DateList` and `PermutationDimensions`. `Min Date` selects the first allocation date on or after `DateFrom`; `DateList` runs from that date through the last allocation date. `PermutationDimensions` adds shifts, sequential periods and roles. Demand already imports that published calendar, so no extra setting or external import is needed.

- `Demand Planning Periods` reads the distinct date/day/shift/period rows before filtering roles. It accepts any positive number of consecutive days and validates one period for each date and source shift.
- `Demand Calendar Prepare` requires `List.Count(#"Demand Roles") * Table.RowCount(#"Demand Planning Periods")` unique role/period cells. Taking the full Settings horizon first prevents missing enabled-role dates from reducing the expected count when other Settings roles still carry those dates.
- The existing Monday start alignment remains: the first planning Monday corresponds to source Week 1. The 14-day source pattern and its three shifts are unchanged; these describe the input grain, not the length of the planning horizon.
- Hours reconciliation follows every full or partial planning fortnight. Partial totals sum the actual included source days and shifts instead of prorating a fortnight average. No output dates or period IDs are generated to fill missing Settings rows.
- The completeness check remains necessary to reject missing or duplicate role/period cells, missing internal dates, inconsistent shared period IDs and incomplete shifts. It no longer imposes a fixed planning length.

Native fixtures cover 1, 7, 14, 15, 21, 28, 35 and 42 days, with unequal source weeks and a sparse zero, plus malformed-calendar cases. The imported calendar is authoritative: these checks cannot detect an intended boundary date omitted from every Settings row or establish whether the saved Settings tables are current.

## Calendar coverage diagnostics and refresh order

The user identified the later calendar-coverage failure as stale `PermutationDimensions` output after Allocation Extraction changed. The required dependency order is to refresh and save Allocation Extraction, refresh and save Settings Data including `PermutationDimensions`, then refresh Demand Extract. Demand reads the saved Settings table; refreshing an upstream query without publishing its updated table leaves downstream imports stale.

`Demand Calendar CHECK` exposes one row per enabled role with expected, actual, distinct, missing and duplicate period counts, plus the affected period IDs. The calendar validation error includes the affected roles and labels available in the saved Settings table. These diagnostics retain the existing completeness rule; they neither invent missing calendar rows nor remove enabled roles. The refresh-order finding does not require weakening the rule.

## External access review

| Import | External source | Path/source handling |
| --- | --- | --- |
| `IMPORT CentriSyncPaths` | Public-machine `CentriSyncPaths.xlsx` | Deliberate fixed bootstrap path; now a named import. |
| `IMPORT Distributed FTE Workbook` | `1. Input/Demand-MasterRoster Manual Read.xlsx` | Derived from `Unit1Path`; shared binary/navigation buffering for allocation, profile and checks. Redundant doubled separators removed. |
| `IMPORT Settings Data` | `2. Calculations/Settings Data.xlsx` | Derived from the same `Unit1Path`; shared settings import. |
| `IMPORT Role Lists` | Existing WhiddonCENTRI SharePoint site | One list-navigation import; `LINK Roles` and `LINK RoleAnalysis` select the original list IDs. This is a SharePoint list connector, not a local file path. |

There are exactly three `File.Contents` calls and one `SharePoint.Tables` call, all in named imports. One canonical `FilePathUrl` lookup feeds the authoritative resolver. It handles CELL filename brackets, local paths, SharePoint paths, Shared Documents, case-insensitive matching and boundary-safe longest-prefix selection. `LocalUnitFolder` remains a simple alias, not a second resolver. The source metadata pathname is repo-relative.

## Validation and remaining prerequisites

- Before the later calendar-diagnostic additions, 66 native Power Query fixtures passed against the edited production queries with synthetic imports and zero external data sources. They covered grouping, effort flags, dynamic role counts, conflicting/missing mappings, shared excluded roles, mixed-inclusion rejection, both supported publication schemas, case/space-normalized DC roles, source zeros and coverage, flexible planning horizons, calendar/timing, full/partial-fortnight reconciliation and resolver cases.
- Seven further diagnostic fixtures bring the suite to 73. Their execution was declined, then the user stopped the investigation and identified the stale Settings calendar. The latest diagnostic changes have not been executed in native Power Query; the earlier 66-pass result does not establish that the current 73-fixture suite passes.
- Nine independent arithmetic/source-contract regression tests passed. Their fixed RN/AIN/AINC4 fixtures are historical test configuration, not the new linked role policy. Their optional saved-workbook mode was not run.
- Shared project validation and the scoped Git whitespace check passed. The shared validator checks configured `Workflows/` files and does not execute M. The native runner targets the adjacent Demand Extract source, not the Settings/MasterRoster changes included in the commit.
- Query headers, final `in` expressions, dependency order, imports and output/check interfaces were reviewed.

Run the native fixtures from the repository root in the signed-in Windows profile:

```powershell
pwsh -NoProfile -File scripts/test-demand-role-mapping-native.ps1
```

The runner uses the existing Power Query SDK test executable and a disposable credential store. It replaces all workbook/SharePoint access with synthetic tables and refuses to execute if an external access expression remains.

Live SharePoint values/credentials, saved upstream publication currency, settings coverage for the actual enabled labels, and the target workbook's `FilePathUrl` input remain unverified. The native fixtures prove the M behaviour against synthetic inputs, not a live Excel refresh. Synchronization and refresh were not performed or authorized by this source edit.
