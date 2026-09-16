# ChatGPT Excel instructions: master-roster versus redistributed FTE

Current Unit1 standard-FTE implementation: [settings source, formulas and validation](2026-09-16-Whiddon-Unit1-Settings-Based-FTE-Conversions.md). Source only; workbook synchronization and refresh remain pending.

Copy everything under "Paste into ChatGPT in Excel" below. These instructions include the query setup, readiness checks, chart design and alignment checks; no separate follow-up clarification is needed.

## Paste into ChatGPT in Excel

Update the existing master-roster FTE comparison charts using the existing worksheet table loaded from `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE` in `Demand-MasterRoster Manual Read.xlsx`. Preserve existing worksheets, charts and filters where possible. The updated query contains both historical and actual redistributed FTE at the same role/week/day/shift grain.

### Query setup and readiness

The revised Power Query source adds two supporting queries:

- `MW FTE Profile Comparison`
- `MW FTE Profile Alignment Check`

Both should be **connection-only**. They do not need separate worksheets or charts. Do not delete or replace any existing worksheet loads if they have already been created; report that situation before changing their load settings.

The role source is now the workbook table `MatchingRosterRoleswithANACCRoles`, with `Roster Roles`, `DC Category`, `DC Role`, and `Direct Care %`. `DC Category = NA` rows are excluded by Power Query; do not delete them from the input table. Chart `Role`/`DC Role` values come from this explicit mapping, not inferred role abbreviations. If the replacement table is missing, stop and report that input/source readiness failure.

`MatchingRosterRoleswithANACCRoles` is an inclusion list. Master Roster roles absent from it are outside the MinuteWorker analysis and are excluded without error. A supplied row with `DC Category = NA` is also excluded. Do not add mappings merely because a non-MinuteWorker role exists in the Master Roster, and do not infer mappings from role names.

If refresh reports `MinuteWorkersFTE publication validation failed`, report the complete error list and stop. Supplied mapping rows must still have unique nonblank `Roster Roles`; RN/OTHER rows require a `DC Role` and numeric `Direct Care %` greater than zero and no greater than 1. Do not bypass publication checks or build charts from stale results.

The chart source remains **`MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE`**. Its existing historical columns are retained, with these added comparison fields:

- `RedistributedRosterFTE`
- `ExpectedRedistributionFactor`
- `RedistributionFactor`
- `ProfileVarianceFTE`
- `ProfileAlignmentStatus`
- `RedistributionMatchCount`

Reuse the existing worksheet table and its query connection. **Do not create a duplicate historical table**, chart directly from the supporting queries, or calculate substitute redistributed values.

Before editing the charts:

1. Check that both supporting queries exist and that the existing historical worksheet table contains the new comparison fields above, together with `HistoricalRosterFTE`, `FortnightWeek`, `Week No`, `Week Day`, `DayOfWeek`, `DayShift`, `Shift`, `ShiftIndex`, `Facility` and `Role`.
2. If the queries are missing, stop and report that the approved revised M source needs synchronization to this workbook. Do not invent replacement queries or change Power Query code.
3. If the queries exist but the worksheet columns are missing or stale, establish whether synchronization and refresh were completed. Do not assume that query existence proves the current code was synchronized. Request the missing synchronization/refresh step rather than using old results.
4. Synchronization and refresh are separate operations requiring explicit approval. These chart instructions do not authorize either. If an approved refresh fails validation, report the error; do not bypass the checks or chart cached results as though they were current.
5. If your Excel tools cannot inspect connection-only queries or confirm readiness, explain the manual check needed. Do not claim that the supporting queries or refreshed results have been verified.

Proceed with the charts only when the updated source table is available and readiness is confirmed.

### Chart purpose and layout

The purpose is to check whether redistribution preserves the historical profile, with only a constant multiplicative scale change. Do not generate the redistributed series by multiplying history; use the actual `RedistributedRosterFTE` column.

For each chart:

- Filter to one `Facility` and one `Role`; begin with BD / RN if no existing selection is available.
- Preserve the existing `Shift` selection and apply the same selection to both measures. When all three shifts are included, keep them as separate positions for a day/shift profile; sum them only for an explicitly daily-total chart.
- Preserve the current chart's day/shift detail. For an all-shifts profile use `DayShift` (MondayAM, MondayPM, MondayNS, through SundayNS), sorted by `DayOfWeek` then `ShiftIndex`: 21 positions per week. For a single-shift chart use `Week Day`, Monday–Sunday, sorted using `DayOfWeek`. Sum shifts into seven daily totals only when the chart is explicitly a daily-total view.
- Values: SUM of `HistoricalRosterFTE` and SUM of `RedistributedRosterFTE`.
- Series: `FortnightWeek` crossed with the two value measures. Keep Week 1 and Week 2 separate, without averaging.
- Exactly four series: **Week 1 — Master**, **Week 1 — Redistributed**, **Week 2 — Master**, **Week 2 — Redistributed**.
- Give Week 1 one colour and Week 2 a contrasting colour. Use **dashed lines for both Master series** and **solid lines for both Redistributed series**. Use markers and straight connecting segments; no smoothing.
- Use a single shared vertical axis starting at zero, labelled **FTE — configured standard roster equivalents**. Do not use a secondary axis, stacking or independent normalisation.
- Keep or add slicers for Facility, Role and Shift. Title the chart with the selected facility, role and shift/all-shifts context.

Use a PivotTable/PivotChart if supported. If per-series line styling is unavailable, use a standard line chart linked to the same four-column summary. Report any styling limitation instead of claiming dashed lines were applied.

### Alignment checks

Before interpreting the chart:

1. Check `HistoricalCoverageStatus`, `HistoricalCellStatus` and `ProfileAlignmentStatus`. Valid profile rows should be PASS or PASS ZERO. Report ERROR, NO TARGET or incomplete coverage; do not hide these by turning nulls into zeros.
2. Retain genuine zero cells. Missing/error points must remain gaps, not fabricated zeros or interpolated points.
3. Display the expected multiplier from `ExpectedRedistributionFactor` once for the selected facility/role. Do not sum repeated factors; confirm its minimum and maximum agree.
4. Show a small alignment summary: the ratio SUM(redistributed FTE) / SUM(historical FTE) separately for Week 1 and Week 2, plus the maximum absolute row-level `ProfileVarianceFTE`. A zero historical denominator means N/A, not zero or infinity. Do not average individual row ratios.
5. The two weekly ratios should match the expected factor when their historical totals are positive. Residuals should be effectively zero (query tolerance: (0.000001 / 60) / ShiftDuration FTE). For a positive factor, profile peaks and troughs should occur in the same positions, even though Week 1 and Week 2 may differ from each other. A valid zero factor produces an all-zero redistributed line.

Keep each role separate for this check. RN and OTHERS can have different scalars; a combined mixed-category profile is not guaranteed to be a constant multiple of history.

Do not divide by standard FTE hours again, apply Direct Care % again, halve/double either series, average corresponding weekdays, or edit the source inputs. The two FTE columns are already at the required units and grain. Do not claim alignment solely because the chart looks similar; report the factor and residual checks.

### Completion summary

Report which existing charts were updated, the facility/role/shift selections, whether all four series and dashed historical lines were applied, and the factor/residual results. State any blocked readiness checks or styling limitations. Do not claim synchronization, refresh or validation was performed unless it actually was.
