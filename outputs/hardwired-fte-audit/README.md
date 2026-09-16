# Hard-coded 7.6-hour FTE audit

Current implementation revision, 16 September 2026: Whiddon Unit1 Manual Read, Demand Extract and Whiddon Cost use the imported `ShiftDuration` directly in hours. The intermediate standard-minutes variable was removed. See the [revised implementation notes](../../docs/2026-09-16-Whiddon-Unit1-Settings-Based-FTE-Conversions.md). Audit snapshots below retain the original pre-change findings.

Reviewed: 16 September 2026. This is a source-text audit, not a workbook inspection or refresh.

## Original audit result (before the settings change)

Eight business M-source files contain literal 7.6-hour or equivalent 456-minute calculations. Four test/diagnostic files and two migration-backup sources also contain these values. The pre-report search snapshot found 214 matching lines across 32 text files; remaining files are documentation, prior audit evidence and run logs. Matching lines include comments and messages, not just arithmetic.

Searched repository working-tree text, including ignored/hidden text, both client trees, both Whiddon units, diagnostics, scripts and backups. Excel files and Excel backup packages, binary report packages/extracts, `.git`, `node_modules` and runtime lock files were excluded. No workbook contents were inspected. The search does not evaluate arbitrary expressions that could produce 7.6 or inspect encoded binary data. Additional M-text searches for a literal 0.95, a 7-hour-36-minute duration and a 0.316666-day representation found no additional matches. Tableau `.twb`/XML text had no literal 7.6 or 456 matches.

Implementation update: the two Whiddon Unit1 demand sources now use imported Settings duration. The original counts, lines and snapshots below are retained as pre-change evidence. The final authorized scope was Whiddon Unit1 and E-O-I only, with Effort-All excluded; Unit2, DATExx, Cost and shared diagnostics were not edited. See [implementation and validation](../../docs/2026-09-16-Whiddon-Unit1-Settings-Based-FTE-Conversions.md). The user later authorized Whiddon Cost, which now also uses Unit1 Settings: see [Cost implementation](../../docs/2026-09-16-Whiddon-Cost-Settings-Based-FTE-Conversion.md). The Cost exclusion above describes the initial scope.

## Business M sources

Prefixes below are repository-relative:

- **W** = `CLIENT/DATExx-Whiddon`.
- **D** = `CLIENT/DATExx`.

| Source file | Arithmetic lines | Hard-coded use |
| --- | --- | --- |
| W `UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m` | 711, 864, 923, 1923, 2121 | Historical hours / 7.6; required roster minutes / 456; FTE-profile tolerance / 456; reconstruct productive hours using FTE x 7.6 x Direct Care %; reconstruct productive minutes using FTE x 456 x Direct Care %. |
| W `UNITS/Unit2/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m` | 705, 858, 917, 1917, 2115 | Same five fixed-denominator calculation/check uses. Unit2 was not changed by the Unit1 net-hours edit. |
| W `UNITS/Unit1/1. Input/2-DemandExtract.xlsx_PowerQuery.m` | 323, 325, 381 | Excluded roster hours = sum(FTE) x 7.6; excluded productive hours = sum(FTE x 7.6 x Direct Care %); DemandHRS = SourceFTE x 7.6. |
| W `UNITS/Unit2/1. Input/2-DemandExtract.xlsx_PowerQuery.m` | 323, 325, 381 | Same three fixed conversions. |
| W `2. Calculations/E-O-I/Effort-All.xlsx_PowerQuery.m` | 231 | `EffectiveAvailability = 7.6 / 8`; applied in capacity and resource-availability branches. Both the standard-hours numerator and eight-hour reference are fixed. |
| D `2. Calculations/E-O-I/Effort-All.xlsx_PowerQuery.m` | 7 | Same fixed `EffectiveAvailability = 7.6 / 8` parameter. |
| W `2. Calculations/Cost/Cost..xlsx_PowerQuery.m` | 19 | `ShiftHrs = 7.6`; standard-hours multiplier for effort-to-cost calculations. |
| D `2. Calculations/Cost/Cost..xlsx_PowerQuery.m` | 19 | Same fixed `ShiftHrs = 7.6` parameter. |

Manual Read and Demand Extract also have comments/messages describing 7.6 or 456. Update those with any approved parameterisation; exact lines are recorded in [matches.json](matches.json).

## Tests and diagnostics

| File | Lines | Treatment |
| --- | --- | --- |
| `scripts/test-minuteworker-period-allocation.mjs` | 80, 85, 100, 114, 121, 123, 132, 149, 154, 185, 245, 251 | Arithmetic fixtures and source-contract assertions assume 7.6/456. Parameterise the applicable calculations and test multiple settings when the source changes; historical 7.6 regression fixtures can remain labelled. |
| `scripts/test-demand-extraction-distributed-fte.py` | 90, 138, 198, 214 | Reference conversions, a deliberately inconsistent shift-duration test and a literal source assertion. Preserve the negative test's purpose when updating the conversion contract. |
| `Workflows/ResidentialCare/Diagnostics/AIN_20260722_HoursTest.pq` | 102 | Tests whether the configured threshold equals 7.6; it does not set the production threshold. |
| `Workflows/ResidentialCare/Diagnostics/AvailabilityDailyWindows_CHECK.pq` | 18, 40, 44 | Supplies a 7.6 allowance in fixed diagnostic scenarios. |

## Historical backup sources

Both contain `EffectiveAvailability = 7.6 / 8` at line 7:

- `CLIENT/DATExx-Whiddon/SyncPathMigration-Backups/20260528-113851/2. Calculations/E-O-I/Effort-All.xlsx_PowerQuery.m`.
- `CLIENT/DATExx/SyncPathMigration-Backups/20260528-113851/2. Calculations/E-O-I/Effort-All.xlsx_PowerQuery.m`.

These are historical references, not inferred edit targets.

## Related files without a literal 7.6/456 conversion

- `1-AllocationExtracted.xlsx_PowerQuery.m` has no literal 7.6 or 456. The Unit1 main output carries `Shift Net Length` as `Hours`; its separate approximate 0.9366 check is a different issue.
- Unit1 `AllocationByShiftAverage.xlsx_PowerQuery.m` reads `ShiftDuration` from Settings for resource FTE and `DurationOfShifts` for role FTE.
- Unit1 `Capacity-ShiftAvailability.xlsx_PowerQuery.m` reads its effective-hours allowance from the Settings `ShiftDuration` value.
- Demand Extract already imports the Settings `ShiftDuration` scalar, but its three conversions listed above still use literal 7.6.

## Evidence and scope

- [All matched files and line numbers](index.md).
- [Exact matched source lines](matches.json).
- [Broader duration/FTE reconciliation](../shift-duration-reconciliation/README.md).
- [Impacted-file register](../../docs/Shift-Hours-FTE-Impacted-Files.md).

No business source or workbook was changed for this audit. Embedded-query parity, worksheet formulas and actual Settings table values remain unverified. Any approved replacement must preserve standard-FTE hours/minutes consistently across publishing and consuming files, and treat the fixed eight-hour factor in Effort-All as a separate business decision.
