# Query structure and import contracts

## Imports and unchanged landing queries

- Always use a separate, meaningfully named `IMPORT` query to read an external source, even when only one extraction is required. Downstream queries must reference that import.
- Use separate `EXTRACT` queries only when more than one distinct extraction is needed from the same imported source.
- NEVER transform or check data in any `IMPORT`, `INPUT`, `EXTRACT`, or `LINK` query. Land the source data unchanged. These queries may only connect/read, navigate to the required source object, and, when useful, buffer the unchanged source or navigation snapshot. An `EXTRACT` returns one source object unchanged.
- Do not filter rows, select or rename columns, change types, clean data, join, calculate, check schemas or data, or perform validation inside any of those four query types.
- Keep approved external source choices and paths unchanged unless path work is requested or separately approved. Use the project's path-resolution instructions for approved path work.

## Downstream work

- Reference the unchanged landed data from meaningfully named downstream queries. Prefer `...Prepare` for preparation work where appropriate; use scope, calculation, or validation queries according to their purpose.
- `...Prepare` is a preference, not a requirement to route every operation through another query. Ask before implementing an unrequested structural change based on that preference.
- Apply client, unit, facility, role, date, or other business-scope filters in preparation or downstream scope queries. Preserve the complete configured key value; do not shorten or derive a partial key unless the approved source contract explicitly requires it.
- For a single extraction, a meaningfully named downstream query may navigate directly to the required source object without an intermediate `EXTRACT` query. For example, `ShiftDuration` can read the `ShiftDuration` table from `IMPORT Settings Data` and select and validate its hours value directly. A separate `...Prepare` query is optional where it helps organize the work. Keep this standard duration in hours.
- Keep existing query/output contracts stable: names, columns, types, row grain, keys, units, and downstream dependencies. Change them only within the requested scope.
- Keep approved or required validation in meaningfully named downstream queries. Do not add data checks merely because the source is being edited, and do not replace required failures with silent empty, null, or default results.

These rules govern query structure and naming. They do not replace the repository's workbook source-of-truth, extraction, synchronization, or refresh requirements.
