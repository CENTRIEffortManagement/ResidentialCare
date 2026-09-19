# Effort-All filter/version audit — 19 September 2026

The current saved workbook is `CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All.xlsx`, outside UNITS. Its embedded M does not contain the Resource = 108 filter. The adjacent `Effort-All.xlsx_PowerQuery.m` does contain it and is stale. The previous answer incorrectly treated this extract as the current workbook state.

## Evidence

Read-only inspection covered 43 disk/history entries: all locally discovered Effort-All workbooks, backups, alternate Effort-AllX and M extracts (including DATExx comparison copies), plus Whiddon Unit1 Effort and its backups and Git versions. Git history was searched across all local refs, with duplicate path/blob combinations inspected once. `audit.json` identifies every snapshot; numbered `.m` files contain extracted evidence, not approved edit/sync sources.

| Version | Embedded/query state | Evidence snapshot |
| --- | --- | --- |
| September 11/13 Git workbook versions | No 108 restriction | 32.m, 35.m |
| September 14 commit 682fb87 workbook | No 108 restriction; Facility = Unit1 in both Availabilities Appended# and Availability Effort | 30.m |
| September 14 commit 682fb87 adjacent source | Git diff introduces Resource = 108; diverges from workbook | Git source diff; same later source blob in 27.m |
| September 17 commit 6df5e61 workbook/source | Workbook still Unit1-filtered with no 108; source still contains 108 | 26.m, 27.m |
| Backup suffix 2026-09-18T20-47-18-337Z (September 19 06:47 Sydney) | No 108; retains Unit1 filters | 3.m |
| Backup suffix 2026-09-18T23-36-56-718Z (September 19 09:36 Sydney) | 108 reappears; former Unit1 filters absent; imports refactored and contract reporting added | 2.m |
| Adjacent extract timestamp September 19 09:41 Sydney | Still 108 | 1.m |
| Current saved workbook, modified September 19 11:38 Sydney | No 108; no Unit1-only predicate in the two availability queries | 5.m |
| September 19 commit 96f42f4 | Commits the current unfiltered workbook alongside the stale filtered M extract | 17.m, 18.m |

Backup suffixes are UTC creation labels, not exact mutation timestamps. Backup LastWriteTime can preserve the source workbook's earlier modified time. These snapshots bracket a change but do not identify its exact execution time or actor.

## Assessment

The evidence establishes temporary reintroduction of the old resource restriction during the interval represented by the two September 19 backups. It is consistent with applying changes from stale M source: the source had retained 108 since September 14 while saved workbooks used different availability logic. It does not establish which tool or person applied it. This was not simply a complete workbook rollback: the affected backup contains new import/refactoring and contract-reporting code alongside the old restriction.

The current workbook has subsequently removed 108 again. Its difference from the adjacent extract is removal of that filter plus repositioning of the AllocationChange import/extract definitions (and omission of extraction metadata). The current embedded section matches the embedded section committed in 96f42f4.

No 108 literal was found in any inspected Whiddon Unit1 Effort M version. This statement concerns Effort, not every upstream workbook in the client.

## Current Effort-All availability behavior

- Availabilities Appended combines both configured facilities without a resource/facility row predicate.
- ResShiftAllocation excludes null Resource rows.
- Availability Effort appends demand, availability and allocation and replaces zero Effort with null.
- ResRoleAvailabilityDevelopedMATRIX pivots this data without another row filter.
- RoleAvailabilityDevelopedMATRIXDELTA groups it without another row filter.
- EffectiveAvailability remains 7.6/8 for the availability branch.
- The historic Facility = Unit1 restrictions are also absent now; facility selection still depends on workbook settings.

No business source or workbook was edited, synchronized, refreshed, saved or closed. The audit reads saved disk content, not unsaved changes in an open Excel instance. Remote/SharePoint revision history was not available in this local audit. Refresh logs do not provide a query-edit provenance trail sufficient to identify the reapplication actor.
