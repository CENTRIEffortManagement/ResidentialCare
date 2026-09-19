# AIN2 expected-difference contract

This contract separates intended correctness changes from changes that must be result-neutral. It is the comparison rule for the corrected three-file, improved three-file, and merged two-file builds.

## Comparison normalisation

- Compare identical output interfaces and declared keys.
- Canonicalise `AIN2` to `AIN` only in Role-bearing fields.
- Require exact equality for text, logical, date, time, duration and integer values.
- Require exact numeric equality except an absolute tolerance of `1e-9` for values whose running-total evaluation order changed.
- Do not canonicalise paths, employee names, IDs, `PreferredRole`, or any non-Role field.

## Approved correctness differences

1. A1 cluster predecessor lookup is scoped by `Resource + PrevIndex`. Rows that previously obtained another Resource's cluster because `ResIndex` restarted may change to their own preceding cluster or to the explicit no-match value.
2. A1 conditional NWD rules now read the explicitly derived three-day allocation fields. Rows affected by the former quoted field-name comparisons may change classification.
3. A1 potential-availability exclusions use `not List.Contains(...)`. Rows that passed only because OR-combined not-equals made the predicate tautological may now be excluded.
4. A2 Role comes from the authoritative Resource x Period skeleton. Rows whose Role previously depended on positional fill-down may change only to the skeleton's role or fail validation.
5. Deterministic Resource/Period tie-breakers may change which equally prioritised row is selected. Every changed key must be attributable to a documented tie; non-tied rows must not change.
6. Missing, extra, duplicate, or wrong-role Resource x Period rows now fail the named validation gates instead of being silently published.

## Changes that must be result-neutral

- Narrowing Cartesian grids and moving column selection/filtering earlier.
- Replacing A2 prefix scans with the linear running-total helper.
- Consolidating repeated A1 workbook navigation.
- Removing unused buffers, dead calculations, duplicate group tables, and redundant rejoins.
- Moving A2 stages into B and publishing the compatibility C#, C## and C### tables from merged B.
- Runner workbook-count and terminal-dependency changes.

These changes must match the corrected three-file build under the comparison normalisation above. Any other row, schema, type, null, error, key, or measure difference is unapproved and blocks acceptance.

## Deliberately unchanged business rules

- Fractional availability behaviour.
- Name-to-EmployeeID migration.
- Existing `Split` branch semantics.
- The freshly verified A1 `ResDayShift` table name and numeric `ID` interface.

Actual changed output keys cannot be enumerated until the separately approved workbook synchronisation, recalculation, refresh, and export run has completed.
