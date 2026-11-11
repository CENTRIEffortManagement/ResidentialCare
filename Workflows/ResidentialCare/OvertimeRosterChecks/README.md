# OvertimeRosterChecks imports-only base

The original-name workbook is a clean base with ten import/path queries and one `UnitPath` sheet containing the `FilePathUrl` named cell. Overtime calculations, business parameter queries, output tables and donor calculation sheets were removed. The workbook was saved, closed and re-extracted; its definitions match the authoritative source. No data refresh was performed.

- Authoritative source: `Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq`.
- Workbook: `CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/OvertimeRosterChecks.xlsx`.
- Matching extracted M: [OvertimeRosterChecks.m](../../../outputs/overtime-roster-checks/bd-reset-20261105/OvertimeRosterChecks.m).
- Move/path receipt: [Overtime-path-receipt.json](../../../outputs/overtime-roster-checks/bd-reset-20261105/Overtime-path-receipt.json).

Retained queries land CentriSyncPaths, published rosters, worker reports, Settings Data, Whiddon SharePoint navigation, role/facility mappings and OvertimeRosterThreshold. The resolver uses the donor's `PathTABLE` and `Unit1Path` style, adapted to `OvertimeRosterChecks.xlsx` under `UNITS/<unit>/2. Calculations`. `FilePathUrl` is recalculated for the new location. The three unit-file imports remain functions taking a unit code; deriving the parent UNITS folder from Unit1Path preserves their existing BD, TE and JH-RY file locations.

The workbook shell originally came from `CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/Worker Reconciliation.xlsx`. The donor remains unchanged.

The preceding attempt is preserved byte-for-byte as `OvertimeRosterChecks-attempt1.pq` in this source folder and `OvertimeRosterChecks-attempt1.xlsx` under `CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/`. Its previous documentation is in [README-attempt1.md](README-attempt1.md). Earlier overtime fixtures and deployment records describe that archived attempt, not this imports-only base.

At the user's subsequent request, BD `Shifts.xlsx` and its authoritative M source were rolled back to their exact pre-overtime versions. The enriched versions were preserved in the rollback evidence folder. See [Shifts-rollback-receipt.json](../../../outputs/overtime-roster-checks/bd-reset-20261105/Shifts-rollback-receipt.json). AllocationByShiftAverage, the Worker Reconciliation donor, Tableau routing and production runner configuration remain unchanged. No data refresh was performed.
