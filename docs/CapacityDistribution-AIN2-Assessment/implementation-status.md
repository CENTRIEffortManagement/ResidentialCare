# AIN2 Capacity Distribution implementation status

Source assessment recorded 19 September 2026. Later workbook status is user-reported and has not been independently verified.

## Completed source work

- Established byte-identical AIN2 workbook copies and freshly extracted their Power Query definitions.
- Preserved attributable baseline, corrected three-file, performance three-file and merged two-file source snapshots.
- Added the explicit `RoleFolder / SourceRole / OutputRole` contract and fail-closed role checks.
- Applied the accepted A1 and A2 correctness and memory changes, including composite predecessor lookup, corrected exclusions, skeleton-owned Role, complete-grid checks, deterministic priority ordering, narrow Cartesian grids and linear running totals.
- Merged A2's period/resource cap stages into B. B has one buffered A1 navigation source, one running-total helper and no A2 workbook path or import.
- Preserved the B query interfaces that load as `ResPeriodAvailabilityCapped_C__TABLE`, `C__TABLE` and `C___TABLE_B` when synchronized to the approved workbook.
- Prepared Capacity routing and role-specific runner profiles without activating production routing or running production jobs.

## Validation evidence

- AIN2 source contracts: 113 passed, 0 failed.
- Native synthetic Power Query fixtures: 10 passed, 0 failed, 0 data sources, 0 remaining PQTest processes.
- Repository Power Query validation: 17 files passed.
- Capacity two-file preparation: 8 passed.
- Batch runner regression: 162 passed.
- Runner launcher regression: 68 passed.
- Duplicate shared definitions: none in A1, A2, B or prepared Capacity source.
- AIN2 workbook lock files: none.
- At the original source assessment, AIN2 workbook binaries matched their copy hashes and no workbook had been synchronized, recalculated, saved or refreshed. The user subsequently reported synchronizing A1 and B, with all resulting tables blank. The exact embedded definitions and current workbook hashes have not yet been independently checked.

## Independent implementation review

The merged B dependency chain was reviewed from its A1 imports through period capping, resource capping, complete-grid validation, C#, C## and C### publication. The review found and corrected one operational issue in the native-test wrapper: PQTest could return results and leave a crash dialog/process behind. The wrapper now preflights its Windows credential store, suppresses crash-dialog UI and removes only processes started by that invocation. A signed-in-profile rerun completed cleanly with no remaining process.

No unresolved static/source defect was found in the reviewed files. The user-reported blank workbook results mean the workbook acceptance gate is currently failed. The source-side A1 imports filter upstream `AIN` through `SourceRole` before relabelling to `AIN2`; the embedded workbook definitions and first empty import must be checked before attributing the blank result to a particular cause. Full correctness, cardinality, runtime and memory acceptance remain unproven.

## Deliberately pending approval

- Once the user has finished other AIN work and closed the AIN2 workbooks, re-extract and compare both user-synchronized definitions with the authoritative source.
- Confirm `FilePathUrl` and the `AIN2 / AIN / AIN2` role tuple, then identify the first empty A1 import or join.
- Correct and separately approve any necessary re-synchronization; then refresh A1 and B and collect the 5/10/20/max cold/warm benchmark matrix.
- Evaluate the runtime and peak-memory acceptance gates.
- Only after those gates pass, move A2 to a timestamped rollback location and proceed to the user-run manual Capacity folder-swap check.

The active AIN2 folder should retain all three workbooks until a populated A1 and the merged B have been verified. No Capacity folder swap or production activation is authorized by this assessment.
