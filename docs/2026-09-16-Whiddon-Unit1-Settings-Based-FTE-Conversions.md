# Whiddon Unit1 — Settings-Based FTE Conversions — 16 September 2026

Revised on 16 September 2026 to use `ShiftDuration` in hours throughout Manual Read, Demand Extract and Whiddon Cost. The earlier standard-minutes layer was removed. This is M-source work only; no workbook synchronization or refresh was performed by this change.

## Approved scope and changed files

The approved scope is Whiddon Unit1, E-O-I with Effort-All excluded, and the explicitly added Whiddon Cost source. The hours-only revision applies to the three sources below.

| Source under Whiddon | Result |
| --- | --- |
| `UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m` | Settings import and validated ShiftDuration; historical/required FTE, reconciliation and tolerance use the setting. |
| `UNITS/Unit1/1. Input/2-DemandExtract.xlsx_PowerQuery.m` | Existing Settings import supplies validated ShiftDuration; demand and exclusion hours use it; saved publication basis is checked. |
| `2. Calculations/Cost/Cost..xlsx_PowerQuery.m` | Cost formulas use ShiftDuration hours directly; ShiftHrs remains a compatibility alias. |
| `2. Calculations/E-O-I/` | No eligible changes: the only literal standard-FTE conversion found is in excluded Effort-All. |

DATExx (including its Cost source), Whiddon Unit2, Effort-All, backups and shared diagnostics were not edited. Supporting tests and documentation were updated. See the separate [dated Cost implementation notes](2026-09-16-Whiddon-Cost-Settings-Based-FTE-Conversion.md) for its import and calendar details.

## Source and formulas

Always use a named IMPORT. Manual Read has one Settings extraction, so it goes directly from `IMPORT Settings Data` to `ShiftDuration`, without an intermediate EXTRACT query. Demand Extract reuses its existing Settings import and retains existing interfaces.

Both read the current unit's `2. Calculations/Settings Data.xlsx`, named table `ShiftDuration`, column `ShiftDuration`. Exactly one numeric, positive, finite hours value is required. Missing, duplicate, text, null, zero, negative and non-finite settings fail. There is no fixed fallback.

```text
ShiftDuration = Settings ShiftDuration hours
HistoricalRosterFTE = historical net roster hours / ShiftDuration
Required FTE = (allocated roster minutes / 60) / ShiftDuration
DemandHRS = published source FTE * ShiftDuration
DemandFTE = DemandHRS / actual DurationOfShifts
```

For example, a 7.6-hour setting stays 7.6 throughout demand and cost conversions. Manual Read's targets remain measured in minutes: it divides allocated roster minutes by 60 before calculating FTE, and multiplies reconstructed productive hours by 60 only when checking minute targets. The standard duration itself is not converted to minutes.

Distribution shares still use productive roster hours. Changing only the standard FTE duration changes the reported standard-FTE quantity, while preserving shares, productive targets, recovered roster hours and actual-shift attendance. Partial and long source rows retain their net hours; no capping or row segmentation was added.

## Publication and workbook prerequisites

- Manual Read retains its approved use of `Shift Net Length`, aliased to `Roster Hours`.
- Manual Read adds the existing flexible path-resolution pattern: `FilePathUrl` must identify that workbook, with one FilePath row or one named cell. CentriSyncPaths must map its location. These workbook prerequisites were not inspected or installed.
- The existing Master Roster import is unchanged. Its previously documented source-location issue remains a separate prerequisite.
- `MinuteWorkersFTE_TABLE` and `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE` each publish `ShiftDuration` in hours instead of the earlier standard-minutes column. Existing fields and row grain are retained; the matrix still selects its original columns.
- Demand Extract requires that basis on every allocation/profile row and compares it to current settings (tolerance 0.0000001 / 60 hours, equivalent to the previous precision). Missing or mismatched values block publication. This catches a duration mismatch; it does not certify general data freshness.
- Demand Extract retains its 12-column output and actual role/shift duration denominator.
- When separately authorized, synchronize both exact sources through the approved process. Manual Read must publish and save the new basis columns before Demand Extract consumes them. Changing the setting later requires regeneration of both paired Manual Read tables before demand consumption.

The profile FTE comparison tolerance is `(0.000001 / 60) / ShiftDuration`; this preserves its former minute-based precision. Both publications and Demand Extract must move together to the hours-based column.

## Validation

- All three revised M sources passed Microsoft's installed Power Query parser.
- 4,394 JavaScript arithmetic/source assertions passed, covering configured durations 7.6, 7.5 and 8, invalid settings, preserved distribution/targets, historical hours, 28-day demand reconstruction and stale saved FTE basis rejection. Manual Read has 52 unique shared queries. Tests also compare the hours-only formulas with the previous calculations and reject saved minute values passed as hours.
- The Python demand fixtures were parameterized, including the optional saved-workbook reader, but were not executed because no Python interpreter was available. The optional workbook-reading mode was not used.
- These checks do not execute M calculations in Excel. No workbook was inspected, synchronized or refreshed.

See the [full impacted-file register](Shift-Hours-FTE-Impacted-Files.md) and [original audit snapshot](../outputs/hardwired-fte-audit/README.md). Files shaded red in the [shift-duration diagram](Shift-Duration-Chart.mmd) identify calculation dependencies, not blanket edit authorization.
