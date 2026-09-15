# Shift duration and effective-hours reconciliation

Reviewed 14 September 2026. This is a static reconciliation of repository text, not a workbook refresh or confirmation of embedded workbook queries.

The code uses several different quantities under similar names: clock duration, net worked hours, a standard 7.6-hour FTE equivalent, average attendance across a shift, and a capacity allowance. There are also three different percentage factors: **0.9366**, **0.936**, and **7.6/8 = 0.95**. They are not interchangeable with the half-hour meal-break factor **7.5/8 = 0.9375**.

## Scope and evidence

Searched working-tree text across `CLIENT/DATExx`, `CLIENT/DATExx-Whiddon` (both units and organisation calculations), `Workflows`, scripts, documentation, existing diagnostic outputs, backups, and Tableau XML. Included hidden and ignored text; excluded `.git`, Excel files and binary packages. No Excel workbook was opened, parsed, compared, edited, synchronized or refreshed. No business code was changed.

The broad search produced 5,169 matching lines in 228 files. These counts include incidental numbers, field declarations, comments and historical artifacts, not 5,169 distinct business rules. The attached [query/file index](index.md), [exact search matches](matches.json), [coverage metadata](coverage.json), and [repeatable collector](collect-evidence.ps1) preserve the search. The collector does not execute M or read referenced workbooks. Query labels in the index are navigation aids based on preceding `shared` declarations.

The `.xlsx_PowerQuery.m` files are existing extracted source text. Their parity with today's Excel workbooks is unverified. In particular, the actual values of the `ShiftDuration`, `Meals`, `ShiftPeriod` and ordinary-work-week workbook tables are not contained in these imports. The user's reported setting of 7.6 hours is consistent with documentation and hardwired conversions, but was not independently read from the settings workbook.

In the references below, **W1** means `CLIENT/DATExx-Whiddon/UNITS/Unit1`; **W2** means the corresponding `Unit2`; **D1** means `CLIENT/DATExx/UNITS/Unit1`. All unit filenames in the table end in `.xlsx_PowerQuery.m`.

## Reconciliation register

| Location and query | Rule and units | Effect and assessment |
| --- | --- | --- |
| W1 `2. Calculations/Settings Data`, `DateNameRoleShiftAllocation`, lines 55–64; W2 same; D1 lines 61–70 | Calculate start/end span in hours; if `> 6`, subtract `.5`. | Separate hardwired break rule feeding `DateRoleShiftAllocation`. Does not consult `Meals`, and never deducts two breaks. The principal allocation output instead comes from AllocationByShiftAverage. Settings' dates/permutations use this branch, but that does not mean its net hours are deducted again downstream. |
| W1/W2/D1 `2. Calculations/Intervals`, `StartTime+Duration`, lines 229–247 | `Duration = EndDateTime - StartDateTime`, converted to number (days). Then compare directly with `MealBreakStart` and subtract `MealBreakTime`. | **Unit mismatch relative to Shifts' use of the same Meals table.** If settings are 6 hours and 0.5 hours, an 8-hour duration is 0.3333 and does not pass `> 6`. Shifts subsequently recalculates net duration from raw `Duration`, so this is not evidence that two breaks currently get deducted. |
| W1/W2/D1 `2. Calculations/Shifts`, `EffectiveDuration`, lines 328–350 | Combine related shifts. If duration `< MealBreakStart/24`, no break; if `> 2*MealBreakStart/24`, subtract two `MealBreakTime/24`; otherwise subtract one. | Active net-work calculation, in days. An eligible 8-hour shift with a 0.5-hour break becomes 7.5 hours. At exactly the first threshold it deducts one break; at exactly twice the threshold it still deducts only one. |
| W1/W2 `2. Calculations/ShiftsRFBI`, `EffectiveDuration`, lines 233–254 | Same parameterized break logic. | Additional variant requiring alignment if the policy changes. Its presence does not establish that it is loaded in the current workflow. |
| W1/W2/D1 `2. Calculations/Shifts`, `DoubleShifts`, lines 353–361 | Assign `RealDuration` and `Effective Duration` to both members of a combined shift. | Carries raw and net duration forward; not another subtraction. `Names_IntervalsList` imports these fields at line 196. |
| W1/W2/D1 `2. Calculations/AllocationByShiftAverage`, `ResourceIntervalAllocation`, lines 168–179 | `ResEffectiveRatio-FTE = Effective Duration / RealDuration` (or zero when `StaffCount = 0`); effective interval effort = ratio × interval duration. | The active percentage reduction in allocation. Spreads a shift's meal-break deduction uniformly over its intervals. For an eligible 8-hour shift, ratio = 0.9375, assuming a 0.5-hour break. This ratio is derived, not hardwired. |
| Same file, `ResourceShiftAllocation`, lines 153–158 | `ResShiftFTE = ResShiftEffort * 24 / IMPORT ShiftDuration`. | Net hours divided by standard FTE hours. At setting 7.6, a 7.5-hour net shift is **0.9868421053 resource FTE**. |
| W1/W2 same file, `RoleShiftAllocation_Calculated`, lines 195–211; D1 `RoleShiftIntervalAllocation`, lines 195–203 | `RoleShiftFTE = RoleShiftEffort * 24 / DurationOfShifts`. | Net hours divided by the role's clock-span hours. At an 8-hour span, the same worker contributes **0.9375 average attendance**. This is a different denominator from resource FTE. W1/W2 validate the role/shift match and hour reconciliation; D1 retains older aggregation/join logic. |
| W1/W2/D1 `2. Calculations/Allocation`, `Allocation` and `ResourceShiftAllocation` | Imports the two allocation outputs. MealBreakTime and MealBreakStart queries are declared but not used to deduct from these outputs. | No additional hardwired meal factor was found in the principal Allocation output. ResourceShiftAllocation applies a configurable full-shift filter; W1's `FullShiftThreshold` is 0. |
| W1/W2 `1. Input/1-AllocationExtracted`, `AllocationExtractedCheck`, lines 127–140; D1 lines 103–116 | Start/end clock hours, summarized to a weekly figure, multiplied by **0.9366**. | Likely the remembered allocation hardwiring. It is in a check query, not the main `AllocationExtracted` query. No downstream M consumer of this check name was found; worksheet/report consumption remains unverified. |
| W1/W2/D1 `2. Calculations/DemandIntervals`, `ShiftDemandUnitINTERVAL`, lines 153–172 | Derive `EffectiveDuration` by subtracting hardwired **0.5** at/above MealBreakStart; calculate `IntervalEffectiveRatio`. | The actual `EffectiveIntervalAttendance` is **DemandFTE alone**. The multiplication by IntervalEffectiveRatio is commented out as inapplicable to ANACC demand. The ratio exists in output but does not reduce the principal demand effort. MealBreakTime is imported but this formula ignores it. |
| Same file, `Table_ShiftDemandHRSCheck`, lines 204–210 | Weekly demand hours × **0.936**. | Independent approximate check, not the principal demand calculation. Differs from both 0.9366 and exact 0.9375. |
| W1/W2 `1. Input/Demand-MasterRoster Manual Read`, `MW Historical WeekDayShift`, line 705; `MW DayShift Allocation`, lines 855–858 | Historical FTE = roster hours / **7.6**. Required roster minutes = productive target minutes / Direct Care %. Required FTE = roster minutes / **456**. | A fixed 7.6-hour source contract. Direct Care % converts productive-care requirement to roster time; it is not documented here as a universal half-hour meal-break rule. Historical input uses `Roster Hours`, not a new start/end-minus-break calculation. |
| W1/W2 `1. Input/2-DemandExtract`, `Distributed FTE Prepare`, line 381; `Demand Extraction Prepare`, line 460 | `DemandHRS = SourceFTE * 7.6`; `DemandFTE = DemandHRS / DurationOfShifts`. | Converts standard FTE equivalents to shift-average attendance. For one source FTE and an 8-hour shift, downstream demand attendance is **0.95**. Checks at lines 323–325 and 472 onward preserve roster and productive hours separately. An imported ShiftDuration query exists but these conversions use literal 7.6. |
| W1/W2/D1 `2. Calculations/Demand`, `ShiftEffort !!` and `ShiftDemandHCAverageANACC`, lines 180–194 and 256–263 | Integrate EffectiveIntervalAttendance × interval days, then divide by summed clock-span days. | Carries average attendance to the role comparison. Does not apply IntervalEffectiveRatio a second time. The older hours/ShiftDuration formula at line 122 is inside a block comment. |
| W1 `2. Calculations/Capacity-ShiftAvailability`, `EXTRACT EffectiveShiftHrs`, lines 96–110; `ResDayShift_Calculated`, lines 385–389; W2 corresponding lines 96 and 383 | Full availability receives Settings ShiftDuration; partial availability receives `min(AvailableHours, allowance)`. The same allowance is the full-day exclusion threshold. | Deliberate availability policy with no additional break deduction. With 7.6 configured, a completely available 7.5-hour window receives 7.6 effective hours. This is documented in `docs/DATExx-Whiddon-Capacity-Shift-Availability-Rules.md`. |
| D1 `2. Calculations/Capacity-ShiftAvailability`, `EffectiveShiftHrs`, lines 38–44 | Ordinary weekly work hours / StdRosterDays × 2, with StdRosterDays = 10. | Older derivation, unlike W1/W2. If ordinary week = 38, this also gives 7.6, but changing the inputs can separate them. W1/W2's retained StdRosterDays values (20 and 10 respectively) do not determine the new effective-hours allowance. |
| W1/W2/D1 role `CapacityDistrib(A.1)-shifts`, `ResPeriodAvailabilityTABLE !!` | Imports EffectiveShiftHrs, then sets `Availability = 1` for each published worker/shift and removes the effective-hours field (W1 AIN lines 28, 83–84). | Partial-shift effective hours determine eligibility upstream but **do not proportionally weight this capacity**. RN and AINC4 repeat the pattern. Existing documentation explicitly defers fractional weighting. |
| `CLIENT/DATExx-Whiddon/2. Calculations/E-O-I/Effort-All`, `EffectiveAvailability`, line 231; consumers lines 245–246, 281. DATExx equivalents lines 7, 118–119, 168 | **7.6/8 = 0.95** multiplies aggregate Capacity, CapacityX and a separate resource-availability Effort branch. | Active hardwired factor, independent of settings and role shift duration. Does not multiply the imported resource allocation. On an 8-hour reference this represents 0.4 hours, not a half-hour break. It may intentionally normalize capacity to standard shift equivalents, but that business intent is not established by the parameter name. The resource branch also retains a Resource=108 filter in W1's organisation source. |
| Both clients `2. Calculations/Cost/Cost.`, `ShiftHrs`, line 19; consumers lines 88, 91, 94 | Hardwired **7.6** multiplies effort measures when calculating costs. | Separate conversion which will not follow a change to Settings ShiftDuration. Verify the units of its effort inputs before replacing the scalar. |
| Both clients `2. Calculations/Change/AllocationChange`, line 10; `Change/SS/ResidualAllocation` and `ResidualAllocation-ALL`, lines 152–153 | Changes divide hours by imported ShiftDuration; residual calculations multiply Demand and Allocation by imported ShiftDuration. | Potential denominator mismatch: role-level Demand/Allocation entering Effort are based on actual shift spans. If these are the fields being converted, multiplying by 7.6 does not reverse a division by 8 or 8.25. These are retained branches; workbook activation and external/manual-input meaning are unverified. |

## What the discrepancies mean

For a hypothetical single worker covering an entire 8-hour shift, with a 0.5-hour break eligible under the configured threshold and standard FTE hours = 7.6:

| Quantity | Result |
| --- | ---: |
| Clock hours | 8 |
| Net hours from Shifts | 7.5 |
| Exact net/clock ratio | 0.9375 |
| Resource FTE, net hours / 7.6 | 0.9868421053 |
| Role average attendance, net hours / 8 | 0.9375 |
| One 7.6-hour source demand FTE, distributed over 8 hours | 0.95 |
| One unit of capacity after Effort-All's 7.6/8 multiplier | 0.95 |
| Eight hours × allocation check factor 0.9366 | 7.4928 hours |
| Eight hours × demand check factor 0.936 | 7.488 hours |

Thus the standard-FTE and average-attendance values differ without necessarily indicating a calculation error. However, **0.95 is not the exact meal-break adjustment**, and the approximate check factors cannot reconcile exactly to the main shift calculation. The 0.9366 check differs from 7.5 hours by 0.432 minutes per 8-hour shift; the 0.936 check differs by 0.72 minutes.

The traced allocation path is `Intervals raw Duration -> Shifts recomputed Effective Duration -> AllocationByShiftAverage effective/raw ratio -> Allocation -> Effort`. Although several files mention EffectiveDuration, the inspected principal allocation path does not demonstrate repeated meal subtraction. The main demand path likewise bypasses its optional meal factor. The material risks are inconsistent units, differing threshold rules, fixed versus configurable denominators, and future changes that accidentally activate another adjustment.

The Intervals mismatch is particularly concrete: its date subtraction yields days, while Shifts explicitly divides the same Meals values by 24. With hour-valued settings, the Intervals effective field is wrong even if current principal allocation recomputation avoids propagating it.

Threshold examples assuming MealBreakStart = 6 and MealBreakTime = 0.5:

| Clock span | Settings TimeWorked | Shifts effective hours | DemandIntervals displayed effective hours |
| --- | ---: | ---: | ---: |
| 6 hours | 6 | 5.5 | 5.5 |
| 8 hours | 7.5 | 7.5 | 7.5 |
| 12 hours | 11.5 | 11.5 | 11.5 |
| 13 hours | 12.5 | 12 | 12.5 |

DemandIntervals' displayed effective hours are not its current attendance multiplier. Intervals' separate days/hours issue is omitted from this table to avoid implying it follows an hour-based rule.

## Additional occurrences and caveats

- `Shifts` line 371 and `ShiftsRFBI` line 275 use 7.5 as a shift-type band boundary, not a net-hours standard. The first band is labelled `0-4Hrs` but tested against `2/24`; that is a separate classification inconsistency.
- `1-AllocationExtracted.Christ`, DATExx `2-DemandExtracted-Master-`, and W2 `2-DemandExtracted-Master-Christ` repeat 0.9366 in check queries. DATExx `2-DemandExtract` line 136 also uses `.9366` for an approximate demand check. W1/W2's newer demand extraction has a different source-FTE contract; do not infer current behaviour from the older variants.
- W1/W2 allocation extraction carries explicit `Break` and net `Hours` from the source roster. Intervals selects/reconstructs start/end duration and Shifts applies the settings break policy. Recomputed net hours can therefore disagree with recorded net hours; the input field names alone do not prove double subtraction. Master-roster history separately uses `Roster Hours`; the actual source meaning requires workbook/source-data verification.
- Demand's importer casts `DemandFTE` to Int64 at line 150 even though the newer extraction publishes fractional attendance. Its principal effort calculation uses the separately preserved numeric `EffectiveIntervalAttendance`, so a direct impact on that main output is not established. Consumers of the rounded field warrant review.
- Demand's `ShiftDemandHCAverageCheck` refers to `ShiftDurations.ShiftDuration` at line 224, while its upstream expansion names the field `ShiftDurations.Duration`. This retained check appears inconsistent with the inspected source and cannot currently be relied upon as proof of hour reconciliation without resolving the field interface.
- Tableau text contains effective-ratio values such as 0.9375 in saved visual state, shift-type labels, and imported duration metadata. The targeted calculation-formula search did not find another direct shift-duration/meal-factor formula. Values such as layout size 450 and EOW ability parameter 0.95 are incidental. Packaged Tableau/binary extracts were not unpacked.
- Existing migration backups and diagnostic M exports contain repetitions of these rules. They are separately labelled in the index and are evidence of past text, not authoritative edit targets.

## Recommended canonical treatment

Define the quantities independently before changing constants: `ClockShiftHours`, `NetWorkedHours`, `StandardFTEHours`, `MealBreakHours`, `MealBreakThresholdHours`, `AverageAttendance`, and `ProductiveCareHours`. Keep the 7.6-hour source-FTE contract explicit; changing it requires synchronized interpretation across source demand, resource FTE, changes and cost.

Use one approved net-work calculation in hours with an explicit threshold boundary and double-shift policy. Convert days to hours at the boundary. Derive the interval weighting ratio from that result. Compare its net hours with recorded roster net hours rather than applying another blanket percentage to already-net hours.

Name or document the two current FTE outputs by their denominators. Use actual role/shift clock hours to reverse average attendance, and standard FTE hours to reverse standard shift equivalents. Decide whether Effort-All's fixed 0.95 is intended as a standard-hours normalization or a meal deduction; if the latter, it does not implement the stated 8-minus-0.5 rule. Capacity's documented full-shift allowance and deferred fractional weighting also need to remain explicit business choices.

Replace approximate check factors with checks that integrate the same net-hour calculation used by the output. Cover no-break shifts, exact threshold, eight hours, exactly twice threshold, longer/double shifts, overnight shifts, and partial availability. Preserve checks that distinguish roster hours from productive-care hours.

No replacement value or policy has been applied. The outstanding evidence for a workbook-level conclusion is the actual Settings tables, embedded-query parity, loaded-query/table mapping, and representative recorded roster hours/breaks. The source search and static reconciliation are complete; these limits do not require guessing that all appearances of 7.6 should become 7.5.
