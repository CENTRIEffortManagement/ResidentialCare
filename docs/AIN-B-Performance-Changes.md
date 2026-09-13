# AIN B performance source change

## Approved target and source verification

The implementation request approved editing only:

`CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AIN/CapacityDistrib(B)-shifts.xlsx_PowerQuery.m`

This is the approved exception to the usual `Workflows/` source location. AINC4, RN, runner code, workbook load settings and matrix construction were not changed. Existing uncommitted source changes were preserved.

Before editing, embedded M was extracted read-only from the exact saved AIN `CapacityDistrib(B)-shifts.xlsx`. The complete section matched the adjacent source after removing its extraction header, normalizing CRLF to LF and trimming terminal whitespace. Internal blank lines and all query text were preserved in the comparison. The verification used stable saved bytes and did not include unsaved Excel edits.

| Evidence | SHA-256 |
| --- | --- |
| Workbook before source editing | `70FA9E65D2711960D957E1B91A4360A4B610B4D48F7F1F940A73250D24E97A70` |
| Approved source before editing | `C603C90A1075E4B8F038C325BF5A4277275CCA15EC36ADA431BC6BAD66748265` |
| Workbook after the user closed the files | `52004E3E894A6C7553E6205C35DEF28EC3DE158A9C02B10E7431A0A6EC9838A7` |
| Final edited M source | `8BE613DDC74EB8CB8316A82594E593FFC3904141EA289B255145FF2E1BCF6447` |

The workbook hash changed when the user closed the files. A second read-only extraction confirmed that its embedded M still matched the original verified definitions. The implementation did not write the workbook.

Detailed evidence and explicitly named review snapshots are under `outputs/ain-b-performance/`. These snapshots are evidence, not alternative edit or synchronization targets.

## Source changes

- `fnAddIndexedRunningTotal` replaces the four per-row prefix scans with a single `List.Generate` traversal. It buffers the supplied ordered table and consumed scalar lists within one invocation, resets at existing group starts, and preserves all-null prefixes, row order, column types and existing cap decisions. It rejects broken positional controls instead of changing grouping or ordering.
- `IMPORTSource A1` provides one buffered file binary and shallow workbook navigation table to the four existing A.1 imports. Their table/sheet selection, types, filters and existing buffers remain intact. Separate loaded output evaluations may still repeat this work.
- The `ResPeriodAvailabilityCapped(C#)TABLE` definition is text-identical to the verified original, including its early buffer and the sequence that filters original nulls, converts zeros to nulls, and then filters role.
- `CapacityDistribB_INPUT_CHECK` returns four small diagnostic rows for the demand, allocation, capped-availability and priority join keys. It reports missing columns, erroneous key cells, missing keys and duplicate key groups. Nullable measures and empty tables with valid key schemas are allowed. It checks the existing consumer inputs; it does not change their filtering.
- `C###TABLE B` explicitly depends on passing input checks. `C###SUM` and `C###MATRIX` retain their existing dependency on that output. Original query names and output columns are preserved.
- The separate correction in `PrioritiseReductionAvail-Setup` now joins Resource to Resource and Period to Period. Its left-outer join and existing priority sort are retained.

The performance-only source was saved as review evidence and passed the running-total fixtures before the join correction and input gate were applied. Its SHA-256 is `65DB7199FD21F73F384F9E6CBDF26C3521FDBDB1146F1BC9C533CCD9E78E73B9`.

New helper/import/check queries should remain connection-only when synchronization is separately authorized. The diagnostic fixture sections are test sources and must not be imported as production workbook queries.

## Validation completed

Microsoft's Power Query parser accepted the edited section. Microsoft Power Query SDK Tools 2.155.2 then executed an isolated synthetic extension containing the actual production helper, input-check logic, corrected join and output gate. The harness substitutes only synthetic inputs and rejects packages containing workbook/file imports.

**115 native M test cases passed, with zero failures and zero data sources accessed:**

- 34 running-total cases: exact totals, original rows, schema and decision outcomes; empty/singleton and multiple groups; nulls, zeros, fractions, cancellation and cap boundaries; tied priorities and invalid positional controls.
- 39 input-check cases: each required key column, null/blank/error keys, duplicate groups, retained null-capacity rows, mixed defects, nullable measures and valid empty schemas.
- Three join cases: distinct Resource/Period pairs, unmatched left rows and nullable priority values.
- 39 publication-gate cases, including a throwing output function that proves failed checks stop output evaluation.

The independent JavaScript arithmetic oracle also passed **942 assertions and 29 source contracts**, including 300 generated group sets, a synthetic 20-day case, all four call-site mappings, original import bodies, the protected buffer/filter query and final-output dependencies.

Native results are recorded in `outputs/ain-b-performance/source-tests-result.json`; the summary is in `outputs/ain-b-performance/validation-summary.json`. The synthetic suites are `Workflows/ResidentialCare/Diagnostics/CapacityDistribB_RunningTotals_TEST.pq` and `Workflows/ResidentialCare/Diagnostics/CapacityDistribB_InputKeys_TEST.pq`. The independent oracle is `scripts/test-capacity-distrib-b.mjs`.

The native fixture runner is:

```powershell
pwsh -NoProfile -File outputs/ain-b-performance/run-native-fixtures.ps1
```

It uses the already downloaded SDK tools in the task's temporary folder and packages only synthetic fixtures. Native M fixture success does not establish actual workbook input quality, Excel refresh speed or runner readiness.

## Deployment and remaining acceptance test

No workbook was synchronized, refreshed, saved or closed by this implementation. No refresh timeout or runner setting changed.

The saved earlier runner attempt failed after its 30-minute deadline while reporting unavailable connection readiness. A separately authorized trial must verify both the results and the runner's completion state.

For that later trial: synchronize only the exact approved workbook/source pair using the approved native mechanism; preserve workbook features and table bindings; re-extract and compare the synchronized definitions. Then validate the exact single-workbook runner selection and perform the explicitly authorized refresh. Measure 5-, 10- and 20-day cases with the existing 30-minute limit, including `C###TABLE B`, `C###SUM` and complete B refresh time. Compare row-level outputs and cap decisions, isolating expected differences from the join correction.

The required successful 20-day Excel refresh and actual-data reconciliation remain unverified. The source implementation and synthetic validation are complete.
