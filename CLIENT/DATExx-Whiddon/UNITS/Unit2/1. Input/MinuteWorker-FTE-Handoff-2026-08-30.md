# MinuteWorker FTE Analysis — Handoff

**Revision:** 8 September 2026 — separate Week 1 and Week 2 allocation.

The user reviewed the two historical weeks and found their differences too large to justify averaging. This revision supersedes the seven-day representative-week calculation. Both historical and required FTE retain all 14 days, with labels such as `1-Tuesday` and `2-Tuesday`.

**Profile-comparison addition:** The existing `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE` now includes actual redistributed FTE beside historical FTE at the same row grain. Its original historical columns remain intact. The addition is source-only, not synced or refreshed. Revised chart instructions are in `docs/MinuteWorker-FTE-Profile-Chart-Instructions.md`.

**Explicit role-mapping addition:** `MinuteWorkersTable` and inferred role abbreviations are superseded by the workbook table `MatchingRosterRoleswithANACCRoles`. Its fields are `Roster Roles`, `DC Category`, `DC Role`, and `Direct Care %`. This addition is also source-only, not synced or refreshed.

## Status and source of truth

- Approved editable source: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx_PowerQuery.m`.
- Target workbook: `CLIENT/DATExx-Whiddon/UNITS/Unit1/1. Input/Demand-MasterRoster Manual Read.xlsx`.
- This revision is **source-only: not synced, refreshed or evaluated in Excel**.
- The user selected this sidecar and the approved IMPORT Extract process. Do not extract older workbook code over the source, substitute a backup, or move the source into `Workflows/` without reconciling that decision with the user.
- The current repository `AGENTS.md` requires explicit opt-in for the named source/workbook pair before synchronization. An edit request does not authorize sync, refresh, save or close.
- Preserve unrelated worktree changes, including the user's changes to `AGENTS.md` and `Master Roster.xlsx`.
- Earlier 7 September code was synced using the approved extension, with backup `Demand-MasterRoster Manual Read.xlsx.backup.2026-09-07T10-19-34-585Z`. That version averaged corresponding weekdays; its sync and any older PASS values do not validate this 14-day revision.

## Goal and confirmed units

Required roster FTE is the final output. Care minutes are an intermediate calculation.

1. `INPUT TargetMinutes` contains productive-care **hours per fortnight**, despite its name.
2. `DC Category` is authoritative: RN receives the RN target, OTHER receives ALL minus RN, and NA is excluded from MinuteWorker calculations.
3. Convert category input hours to fortnight minutes by multiplying by 60.
4. Historical productive hours = roster hours × configured Direct Care %.
5. Historical FTE = roster hours / 7.6, without applying target scaling or dividing by Direct Care %.
6. Required roster minutes = allocated productive minutes / Direct Care %.
7. Required FTE = required roster minutes / 456. These are 7.6-hour shift equivalents, not full-time employee positions.
8. Keep full precision; round only for presentation.

## Fortnight identity and ordering

- `Week No`: original source identifier, retained unchanged.
- `FortnightWeek`: 1 or 2. `MW Historical Weeks` maps the two source week numbers in ascending numeric order within each facility.
- `DayOfWeek`: Monday = 1 through Sunday = 7.
- `FortnightDay`: `1-Monday`, `1-Tuesday`, …, `2-Sunday`.
- `FortnightDayIndex`: 1 through 14; use this field to sort charts and tables.
- `FortnightDayShift`: e.g. `1-Tuesday-AM` in historical staging.
- `WeekdayShift`: now the same week-qualified day/shift style in the allocation and matrix.

Exactly two complete source weeks are required for each target facility. A one-week or longer extract fails validation; no pair is silently selected or averaged. All supplied historical weeks remain visible diagnostically. If numeric week order does not represent chronology, for example a year-boundary extract with weeks 52 and 1, obtain an explicit period mapping before use. A repeated week number across years cannot be distinguished without a better source period key.

Absent role/shift cells in complete facility weeks are explicit zero cells. Missing cells in incomplete weeks and cells with null/negative source hours remain flagged/null. Seven-day coverage is inferred from eligible source rows; it does not prove that every employee was present in the source extract.

## Allocation formula

For each facility/category and distinct source week/day/role/shift:

```text
CellProductiveHours = HistoricalRosterHours × DirectCarePercent
FortnightShare = CellProductiveHours / SUM(CellProductiveHours across all 14 days)
AllocatedProductiveMinutes = CategoryFortnightHours × 60 × FortnightShare
RequiredRosterFTE = AllocatedProductiveMinutes / DirectCarePercent / 456
```

There is **no weekday averaging, division by two, or doubling of results**.

The equivalent two-stage calculation uses that individual fortnight day's category share, then its role share, then the role/day's shift share. Both methods are checked for agreement.

`RoleWeeklyTargetMinutes` now refers to the actual subtotal for the row's `FortnightWeek`; Week 1 and Week 2 may differ. `RoleTargetMinutes` sums all 14 days. `RoleAverageDailyTargetMinutes = RoleTargetMinutes / 14` and `CategoryDailyTargetMinutes` remain informational fields only; neither drives allocation. Do not sum repeated target fields across shift/day rows.

## Query pipeline

New/supporting staging:

- `INPUT MinuteWorkers`: compatibility query name retained; now returns `MatchingRosterRoleswithANACCRoles` directly without an assigned-types step.
- `MW Roster Role Mapping Prepare`: normalises explicit roster and DC-role keys and maps DC Category to internal MinuteCategory.
- `MW MinuteWorkers Prepare`: one retained configuration per normalised DC Role; many roster roles may share one DC Role.
- `MW Fortnight Days`: fixed ordered 14-day display keys.
- `MW Historical Weeks`: original source-week mapping and period count.
- `MW Historical WeekDayShift`: unaveraged, zero-completed historical cells.
- `MW Historical DayShift`: positive-history cells with conditional day shares and full-fortnight shares; no averaging.
- `MW Role Distribution`: role share within each individual fortnight day.
- `MW Category Weekday Targets`: name retained, but now 14 weighted category/day targets.
- `MW Role Targets`: individual day targets, actual weekly subtotals and fortnight totals.
- `MW DayShift Allocation`: required FTE at facility/role/FortnightDay/shift.
- `MW FTE Profile Comparison`: exact role/week/day/shift join to actual allocation, with independent expected-scalar and residual checks.
- `MW Role Fortnight Day Totals`: sums shifts into explicit 14-day role totals.

Published outputs:

| Query | Current grain/use |
| --- | --- |
| `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE` | Paired historical and redistributed FTE per original week/day/role/shift, with scalar-alignment diagnostics; publication-gated |
| `MinuteWorkersFTE_TABLE` | Validated required FTE per distinct fortnight day and shift |
| `MinuteWorkersFTE_MATRIX` | One row per facility/category/role; separate week-qualified day/shift columns, up to 42 |
| `MinuteWorkersFTE_WEEKLY_DISTRIBUTION_CHECK` | 14 rows per allocated role; individual day, actual week and full-fortnight reconciliation |
| `MinuteWorkersFTE_CATEGORY_DAILY_CHECK` | 14 rows per facility/category, reconciling weighted day targets and both weeks |
| `MinuteWorkersFTE_HISTORICAL_DAY_CHECK` | Historical versus required FTE for each distinct role/week/day, with source coverage and headcount context |
| `MinuteWorkersFTE_CHECK` | Main input, history, distribution and original-hours reconciliation results |

Legacy `LocRole…` query definitions remain in place, but their shared `Master Prepare` source now uses the explicit mapping and excludes NA roles. They still include the old historical-hours averages; do not use them to infer the new 14-day target allocation.

## Changed interfaces for existing charts

- Use `FortnightDay` sorted by `FortnightDayIndex`, not `Week Day` alone, for a 14-day axis.
- Existing matrix column names change to week-qualified labels such as `1-Tuesday-AM`.
- The seven-row proof tables now have fourteen rows per role/category. Weekly subtotal fields repeat within their own week; do not sum repeated subtotal fields.
- The historical daily comparison retires the old `HistoricalAverage…`, across-weeks and average-versus-allocator fields. Use `HistoricalDailyRosterFTE`, `HistoricalAMRosterFTE`, `HistoricalPMRosterFTE`, `HistoricalNSRosterFTE`, `TargetDailyFTEShiftTotal`, and `TargetVsHistoricalRosterFTE`.
- Direct master-roster headcount context is retained as `HistoricalDailyDistinctWorkers` and `HistoricalRowsMissingEmployeeCode`; no averaging is applied to these fields.
- Whole-period shares are explicitly named `CategoryFortnightRoleDayShiftDistribution%`, `CategoryFortnightDayDistribution%`, and `CategoryFortnightHistoricalProductiveHours`.
- Diagnostic tables may be inspected independently, but their existence is not proof that publication checks pass. Historical coverage/cell flags must be reviewed.

## Validation and acceptance

Required checks:

1. Every supplied normalised `Roster Roles` key is nonblank and unique. Unmapped Master Roster roles are outside the MinuteWorker inclusion list and are excluded without error.
2. `DC Category` is RN, OTHER, or NA. Retained RN/OTHER mappings have a DC Role and `0 < Direct Care % ≤ 1`; mappings sharing a DC Role have one category and percentage.
3. Raw roster rows have non-null, non-negative hours and valid week/day/shift keys.
4. Every target facility has exactly two source weeks, each with seven eligible weekdays.
5. Historical graph cells are unique; source row counts and unaveraged roster hours reconcile.
6. Conditional role/day and role/day/shift shares reconcile within each individual fortnight day.
7. Full-fortnight category cell shares sum to 100%.
8. Positive category targets require eligible productive history somewhere in the fortnight. Zero-history days may receive zero.
9. Individual day allocations, the two actual weekly subtotals and the full fortnight reconcile.
10. Independently, `sum(FTE × 7.6 × Direct Care %)` over **all 14 days once** equals the original RN and ALL input hours. There is no factor of two.

11. `MW FTE Profile Alignment Check` checks every historical cell against actual redistributed FTE. Expected factor = category target productive minutes / 60 / historical category productive hours. Because direct-care adjustment cancels on conversion back to roster FTE, the same factor applies to every role, shift and day in that facility/category, across both weeks. RN and OTHERS can have different factors. A mixed-category aggregate is not guaranteed to retain one scaled shape.

Failures block `MinuteWorkersFTE_TABLE`, `MinuteWorkersFTE_MATRIX` and `MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE` through `MW Publication Check`. Inspect connection-only `MW Historical WeekDayShift` for raw history when the paired output is blocked.

Profile comparison fields:

- `HistoricalRosterFTE`: unchanged, unaveraged master-roster hours / 7.6.
- `RedistributedRosterFTE`: actual joined FTE from `MW DayShift Allocation`.
- `ExpectedRedistributionFactor`: expected category-wide multiplier; repeated informational field, never sum it.
- `RedistributionFactor`: redistributed / historical FTE for positive historical cells; null for zeros.
- `ProfileVarianceFTE`: actual redistributed FTE minus historical FTE times expected factor; expected zero.
- `ProfileAlignmentStatus`: `PASS`, `PASS ZERO`, `ERROR`, or `NO TARGET`.
- `RedistributionMatchCount`: exposes missing or duplicate allocation joins without multiplying historical rows.

Only valid zero-history cells with a configured target can receive an absent-allocation zero. Missing targets, incomplete history or missing positive-history allocations do not become fabricated zeros. An all-zero target/history pair can pass with an undefined ratio; no unique scalar can be inferred from zero divided by zero.

Current source verification:

- `scripts/test-minuteworker-period-allocation.mjs`: 544 independent arithmetic/source-contract assertions passed, including balanced source delimiters, the 14-day/profile cases, explicit overrides of former guesses, many-to-one DC roles, NA/unmapped-role exclusion, duplicate mappings, invalid/non-numeric percentages, invalid categories and conflicting DC-role attributes.
- 47 uniquely named shared queries.
- Lexical delimiter/string checks passed; the quoted shared-query dependency scan found no cycles. These are not a full M parser.
- The shared source validator passed six M files using a temporary configuration including this sidecar's directory. That configuration was removed. The validator checks nonempty files and conflict markers, not runtime calculation.
- The selected sidecar's input and historical-preparation prefix was deliberately revised for the explicit role mapping; unrelated legacy calculations were preserved. Query headers, mapping dependency flow, joins, final outputs and the week-qualified matrix key were reviewed. `git diff --check` passed apart from line-ending warnings.
- These fixtures do not execute Power Query. Excel runtime validation remains pending; do not describe source-only tests as refreshed workbook results.

Next authorized workflow: obtain explicit sync approval, validate the named source, confirm the named workbook is closed, sync via the approved mechanism, and re-extract to a separate comparison artifact as required by AGENTS.md. Never overwrite this approved source during verification. Obtain refresh approval separately, then inspect the check outputs before using the matrix.

## Plotting

For a continuous fortnight comparison, use `FortnightDay` on the axis sorted by `FortnightDayIndex`; filter facility, role and shift, or sum shifts for daily totals.

For two historical-week series, use `Week Day` on the axis with `FortnightWeek` as the series. This compares them visually without averaging.

For the requested four-line overlay, use the same weekday axis and combine `FortnightWeek` with the two value measures `HistoricalRosterFTE` and `RedistributedRosterFTE`. Use one colour per week, historical dashed and redistributed solid, with a common FTE axis. Filter to one facility and role; filter shift or sum all three for daily totals. Never normalise the two profiles separately, since that would conceal the scalar difference. Check `ProfileAlignmentStatus` and the residual before claiming alignment.

For historical versus required daily FTE, use `MinuteWorkersFTE_HISTORICAL_DAY_CHECK` with `HistoricalDailyRosterFTE` and `TargetDailyFTEShiftTotal`. Review coverage flags and publication checks first. Missing/error historical values must not become fabricated zeros.

## Explicit role mapping

`MatchingRosterRoleswithANACCRoles` is the sole authority for roster-role relationships. Matching is case-insensitive after trimming, but no substring, abbreviation, nursing qualification, or role-name inference is applied.

- `Roster Roles` is the original role label from `IMPORT Master`.
- `DC Role` is the assigned analytical role. Its normalised uppercase key is used for grouping so case-only differences cannot split a role.
- `DC Category` is RN, OTHER, or NA. Internal `MinuteCategory` remains RN/OTHERS for compatibility with target calculations.
- `Direct Care %` converts historical roster hours to productive hours and target productive minutes back to roster FTE.
- Multiple Roster Roles may map to one DC Role and are aggregated only after the explicit join.
- NA rows remain in mapping preparation for validation but are removed from `Master Prepare` before analysis; no Excel table rows are deleted.

The old `MW Abbreviate Role`, `QFR Category`, and guessed mappings are retired. `MinuteWorkerRoleAssignments_TABLE` exposes the retained original-to-DC-role assignments. Mapping-table roles with no retained historical rows remain warnings; Master Roster roles absent from the mapping table are excluded from this analysis.

## Last observed target values

The original input values below were recorded on the snapshot date. Their units were corrected on 7 September; these are calculated expectations, not results of a refreshed workbook. Reconfirm the raw inputs before using them as regression fixtures.

| Facility | RN input hours/fortnight | ALL input hours/fortnight | OTHERS hours/fortnight | RN minutes/fortnight | ALL minutes/fortnight | OTHERS minutes/fortnight |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| BD | 1041.117 | 5180.46666666667 | 4139.34966666667 | 62467.02 | 310828 | 248360.98 |
| JE | 395 | 1954 | 1559 | 23700 | 117240 | 93540 |
| TE | 738 | 3641 | 2903 | 44280 | 218460 | 174180 |

For BD RN at 100% direct care, 1041.117 hours per fortnight gives approximately 136.989078947 roster shift equivalents across all 14 days. The two weekly subtotals may differ; their mean is approximately 68.494539474. These are arithmetic expectations, not refreshed results of this revision.

The facility code is `TE`; an older `NR`/`TE` mismatch was reported as resolved.

## Last observed historical role coverage

Counts below are diagnostic row counts after role normalization, not target amounts.

| Facility | AIN | CCCM | CSM | EN | RN | RSM | WCSO |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| BD | 446 | 10 | 10 | 112 | 114 | 10 | 30 |
| JE | 154 | 9 | 10 | 28 | 40 | 10 | 14 |
| TE | 323 | 7 | 20 | 25 | 83 | 20 | 43 |

## Remaining safeguards

- Historical data means master-roster hours, not necessarily hours actually worked.
- CAM/RGM unmatched-history warnings do not authorize invented allocation weights.
- Shift boundaries still depend on input row order. This was not changed in the 14-day revision.
- Legacy outputs still share the configured-role preparation filter; they are not independent of future configuration changes.
- Preserve the hours-per-fortnight correction and approved source path. Do not roll back wholesale to an older allocator.
