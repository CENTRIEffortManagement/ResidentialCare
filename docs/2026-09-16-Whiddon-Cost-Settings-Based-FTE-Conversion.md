# Whiddon Cost — Settings-Based FTE Conversion — 16 September 2026

Revised on 16 September 2026: `ShiftDuration` remains in hours, and all Cost external imports now use the standard FilePathUrl/CentriSyncPaths resolver. The legacy Folder input and intermediate standard-minutes query have been removed.

The user explicitly expanded the approved scope to Whiddon's Cost source after the Unit1 demand changes. DATExx Cost and Effort-All remain unchanged.

## Changed source

`CLIENT/DATExx-Whiddon/2. Calculations/Cost/Cost..xlsx_PowerQuery.m`

- `IMPORT CentriSyncPaths` reads the public-machine mapping table. `CostPathTABLE` resolves this workbook's `FilePathUrl` using the standard case-insensitive, boundary-safe longest-prefix match for local and SharePoint roots, including Shared Documents URLs.
- `FilePathUrl` must contain one FilePath row or a single named cell exposed as Column1. Excel CELL filename brackets and sheet suffixes are handled; the filename must be Cost..xlsx. `ClientPath` removes the expected `2. Calculations/Cost` suffix and verifies the DATExx-Whiddon client root. Invalid or unmapped paths fail visibly.
- `IMPORT Settings Data` reads `ClientPath` plus `UNITS/Unit1/2. Calculations/Settings Data.xlsx`.
- `IMPORT Inefficiencies` reads `ClientPath` plus `2. Calculations/E-O-I/Inefficiencies.xlsx`. Its one table extraction remains directly in EffortOutcomesAG1_1DayShiftAB; no intermediate EXTRACT query is needed.
- No legacy Folder input or competing path resolver remains. The only fixed bootstrap path is the public-machine CentriSyncPaths workbook. The source provenance header now uses the correct repo-relative Whiddon Cost path.
- Settings supplies two extractions: `EXTRACT ShiftDuration` and `EXTRACT PermutationDimensions`.
- `ShiftDuration` validates and returns exactly one numeric, positive, finite hours value directly. There is no minute conversion or fixed fallback.
- Cost formulas now multiply by `ShiftDuration` directly. `ShiftHrs = ShiftDuration` remains solely as a compatibility alias with the existing parameter metadata.
- `PermutationDimensions` now aliases its extraction from the same Unit1 Settings import. This replaces the legacy `FACILITIES/ASHB` settings path.

## Effect on cost

The existing unnecessary-allocation, shortage-replacement and overtime formulas multiply directly by `ShiftDuration`, Unit1's configured standard-FTE hours. Rates, existing zero overtime multiplier, role/facility exceptions, output columns and summaries were preserved.

Unit1 is the common settings authority for this organisation Cost model; no per-facility duration mapping was introduced. This change parameterizes the existing multiplier. It does not establish that every incoming effort measure has a standard-FTE denominator; the earlier measure-basis review remains relevant. The Inefficiencies table selection and transformations are unchanged; its file access now uses the standard resolver.

## Validation and workbook status

- Microsoft Power Query parser: passed.
- The intermediate standard-minutes query was replaced by ShiftDuration; the original Cost interfaces remain. There are 20 unique shared queries after adding the resolver and Inefficiencies import.
- 33 Cost path fixture/source checks passed: local paths, named-cell CELL output, forward slashes, SharePoint, Shared Documents, UNC roots, case-insensitive matching, longest-prefix selection, and rejection of malformed/unmapped/wrong-workbook/wrong-client paths. These are independent reference checks, not M execution.
- All 4,394 FTE arithmetic/source assertions passed. The path-only revision leaves the cost formulas and imported table transformations byte-for-byte unchanged.
- Source comparison confirmed the only cost-formula change was substituting ShiftDuration for ShiftHrs; rates and business rules were preserved. Arithmetic fixtures confirmed equivalent costs at 7.6, 7.5 and 8 hours.
- No literal 7.6-hour/456-minute setting or legacy ASHB Settings path remains in this Cost source.
- Query headers, dependency order, final results and output interfaces reviewed; diff whitespace check passed.
- No workbook was inspected, synchronized or refreshed by this source change. The FilePathUrl input, public mapping, Settings tables and runtime outputs remain to be verified when workbook operations are authorized. A missing FilePathUrl requires separately authorized workbook setup; there is no fallback to Folder.

## Reinforced instructions

AGENTS.md now requires proactive review of every external import in an approved edited M source. Apply the standard resolver, remove legacy path inputs within that source, and verify every file access before reporting completion. Preserve approved source choices and keep workbook operations opt-in. The original three-workbook worksheet-copy procedure does not automatically authorize Cost worksheet changes.

See [Unit1 demand implementation](2026-09-16-Whiddon-Unit1-Settings-Based-FTE-Conversions.md) and the [full impacted-file register](Shift-Hours-FTE-Impacted-Files.md).
