Post-run review, 29 October 2026 (Australia/Sydney).

Cost and EOW were excluded at the user's request. This review covers the latest receipts for 87 other workbooks across BD, TE, JH-RY, organisation reporting and Leave Balance Stats. It combines the completed jobs from run `20261029-074441-a2fe3870` with the four completed jobs from `20261029-085859-fbb9b367`. The latter ran from 08:58:59 to 09:00:47 AEDT, exit 0.

The runner tests passed, but the saved workbook results are not clean. The most significant findings are the Excel errors in Inefficiencies, skipped allocation records, and loss of freshness after later TE file changes.

| Finding | Saved evidence | Implication |
| --- | --- | --- |
| Errors inside Inefficiencies calculation tables | `2. Calculations/E-O-I/Inefficiencies.xlsx`: `EffortOutcomesAG1_1DayShiftEF`, sheet `3Pool-EF-AG1.1`, range `A1:BM1021`, contains 3,189 error cells (3,123 `#VALUE!`, 66 `#DIV/0!`). `EffortOutcomesAG1_1DayShiftIN`, sheet `3Pool-IN-AG1.1`, range `A1:BS1021`, contains 19,737 `#VALUE!` cells. | 22,926 saved errors are inside these two calculation tables. Successful query refresh and save did not reject these worksheet errors. Their effect on published reporting values has not been independently reconciled. |
| Skipped allocation records | A.1 `CapacityDiagnostics_SUMMARY`: BD/CS reports 76, BD/EN 65, JH-RY/AIN 104 skipped records. Their allocation-name-to-Resource checks report 2, 1 and 1 failures respectively. | 245 allocation records were skipped in these role calculations. These are diagnostic counts, not 245 distinct workers. |
| Final cap breaches | `BFinalCap_CHECK`: BD/AIN 34, BD/CS 26, JH-RY/CS 12, TE/CS 9 failed results. | All 81 failed results have `Origin = Preexisting allocations` and `OptionalSlots = 0`. The saved evidence attributes these breaches to existing allocations rather than added optional spare. |
| Effort-All facility check fails | `2. Calculations/E-O-I/Effort-All.xlsx`, `ResourceContractOverage_CHECK!A1:D5`, row 2: `Facility1 resolves to Unit1`, Fail, Details `Facility1 resolves to BD`. The same result appears in `ResourceContractOverage!A1:D5`. | The saved overage check expects Unit1 while the configured facility resolves to BD. This is one repeated condition, not two separate failures. Whether the check should apply to BD requires a business-scope decision. |
| Later TE saves invalidate run freshness | TE `Capacity-ShiftAvailability.xlsx` changed at 09:13:44 AEDT; TE `Capacity.xlsx` changed at 09:33:04 AEDT. Both differ from their recorded run receipts. | The final derived assessment has 66 Verified outputs, 2 ChangedAfterRun outputs and 19 InputsChanged outputs. The changed inputs affect TE Staff, all twelve TE role-stage workbooks, TE Effort, organisation StaffAll/EffortAll/Outcomes/Inefficiencies, and Tableau Connection. Later saves are not adopted as runner refresh receipts. |

The latest four-file `-CheckOutputs` command still returns exit 0. Its own selected outputs and declared file stamps match. The combined assessment also checks the upstream producers from the earlier run and exposes the TE freshness issue. Leave Balance Stats has ordering dependencies but no declared data-input inventory in this plan, so its verified receipts do not establish complete input lineage or business reconciliation.

Additional saved findings:

- Existing-allocation five-worked-day checks fail in BD AIN/CS/EN/RN, JH-RY AIN/CS/EN/RN, and TE AIN/RN. A.2 reports the corresponding saved eligibility/calendar breaches. These repeated stage observations must not be added together as distinct workers.
- The shared demand master has 48 `#REF!` formula results on `BD RN Historical FTE`, including `B43`, `B44` and `B46`. These are outside Excel tables. They do not establish that the published demand tables failed.
- Organisation StaffAll, Outcomes, Inefficiencies and Tableau Connection each have three `#VALUE!` cells in FILEPATH helper sheets, at `B15`, `C15` and `C20`. Their effect on current imports was not established.
- TE `ResourceContract_CHECK` has ten Pass rows and one Warning for a worker using the Settings roster-cap fallback. BD and JH-RY each have eleven Pass rows.
- Demand-master diagnostics report 22 roster-location warnings for locations outside configured Units. The same warnings repeat in several diagnostic views; they are not 88 different locations. The inspected demand extraction/calendar, CS allocation, historical coverage, daily allocation, FTE profile alignment, Availability rules, worker reconciliation and ResDayShift status/severity checks otherwise report passing results.
- 84 workbooks contain defined names with `#REF!`. This is a structural inventory finding; active use and business impact were not established. No names were removed.

The saved values were read from 810 Excel tables, including 121 tables identified as checks or diagnostics. Every worksheet was scanned for saved Excel error cells. Diagnostic tables were reviewed using their Status, Severity and named status columns; row-only diagnostics remain evidence rather than implicit passes. Twelve B availability-lineage issue tables are empty. Additional check-like connections in 49 workbooks have no worksheet query-table load surface; their individual results were not independently evaluated. No M source was edited or extracted, and no query was evaluated afresh.

Current access validation returned exit 1 with twelve target/input failures. Tableau holds BD Shifts, AllocationByShiftAverage, Shift-StaffDistribution and Effort, organisation Tableau Connection, and Leave Balance TableauConnect-LB. Excel lock files are present for TE Capacity, TE Effort and organisation Effort-All. Some failures repeat because a locked output is also another job's input. These are current rerun blockers, not evidence of failed saves in the completed run. No applications were closed or processes terminated by this review. The initial sandbox Restart Manager error was resolved by the approved read-only validation outside the sandbox.

Nine regression suites passed without Excel:

| Suite | Result |
| --- | --- |
| `scripts/test-batch-refresh-runner.ps1` | 413 assertions passed |
| `scripts/test-excel-refresh-safety.ps1` | Readiness, loaded-refresh evidence, required checks, save/close receipts, timeout, stop and cleanup cases passed |
| `scripts/test-batch-output-status.ps1` | 75 assertions passed |
| `scripts/test-batch-output-status-cli.ps1` | 27 checks passed |
| `scripts/test-batch-runner-recovery.ps1` | 55 checks passed, including deliberate Windows lock/access failures |
| `scripts/test-runner-file-access.ps1` | 26 checks passed |
| `scripts/test-git-aware-refresh.ps1` | 52 checks passed; optional live-process checks were not run |
| `scripts/test-excel-save-diagnostics.ps1` | 29 checks passed |
| `scripts/test-batch-runner-launcher.ps1` | 133 checks passed |

Excel was not launched by this review. No workbook was refreshed, saved, synchronized or closed. SHA-256 checks confirmed each workbook was stable during its read. TE Capacity changed after its first inspection and was reinspected from its newer saved version. Both run ledgers were unchanged during the full inspection. These observations do not establish when the saved diagnostic defects first arose, nor include unsaved changes in already-open workbooks.

Priorities for a separately requested repair are to investigate the two Inefficiencies calculation tables, resolve the three role allocation-identity failures, and reconcile the TE chain after the later saves. The pre-existing cap/calendar breaches and Effort-All facility scope also need review. No repairs or refreshes were made.

Evidence files in this folder:

- [Saved tables, diagnostic rows, exact locations, cell errors and hashes](saved-checks.json).
- [Combined metadata assessment using latest receipts per job](combined-metadata.json). This is derived audit evidence, not a replacement run ledger.
- [Per-run metadata assessment](run-metadata.json).
- [All inspected Status and Severity counts](all-check-statuses.json).
- [Check-table register](check-table-register.json).
- [Cell-error locations and formula samples](triage.json).
- [Current access-validation output](production-validation.log).

Workbook paths above are relative to `CLIENT/DATExx-Whiddon`; role workbooks are the exact A.1/A.2/B targets named in [the inspected selection](selected-workbooks.json). Original logs remain under `CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh/20261029-074441-a2fe3870` and `CLIENT/DATExx-Whiddon/RunLogs/BatchRefresh/20261029-085859-fbb9b367`.
