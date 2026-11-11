# ExcessSegments position-parts diagnostic and conditional proposal

Proposal only, 5 November 2026, Australia/Sydney. No M source or workbook changes are made by this document. The reviewed canonical source is `Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq`, SHA-256 `23F8DA4322DB02598D9581B743FE16C412A52407412AB98DDF7C5F00335AF882`.

The current full-refresh attempt has no completed output at the supervisor's eight-minute report. Its reported private-memory range of 8–11 GB is below the stopped attempts' pressure, but does not establish successful completion or acceptable elapsed time. Do not apply a further change while that attempt is running.

## Smallest next diagnostic

If the current attempt times out or is stopped, first probe the existing `ExpandedParts` step at line 684. Force every scalar cell in its sixteen columns, not only its row count. The current inputs are eleven shift fields plus five position fields, and the nested `Parts` column has been removed. Capture elapsed time, row count, memory and unchanged-source/workbook evidence using the already approved bounded probe process.

If that step completes within the bound, the smallest subsequent production proposal is the following single buffer and two reference replacements. It preserves all current grouping, join and fallback construction:

```diff
     ExpandedParts = Table.RemoveColumns(ExpandedPositionParts, {"Parts"}),
+    BufferedExpandedParts = Table.Buffer(ExpandedParts),
     WorkSegments = Table.Buffer(OvertimeRosterChecks_WorkSegments),
-    WorkJoin = Table.NestedJoin(ExpandedParts, {"SourceShiftID"}, WorkSegments, {"SourceShiftID"}, "ShiftWork", JoinKind.LeftOuter),
+    WorkJoin = Table.NestedJoin(BufferedExpandedParts, {"SourceShiftID"}, WorkSegments, {"SourceShiftID"}, "ShiftWork", JoinKind.LeftOuter),
     // Existing SegmentLists calculation remains unchanged.
-    CombinedSegments = if Table.IsEmpty(ExpandedParts) then {} else List.Combine(SegmentLists[ExcessSegments]),
+    CombinedSegments = if Table.IsEmpty(BufferedExpandedParts) then {} else List.Combine(SegmentLists[ExcessSegments]),
```

Both references must change; leaving the emptiness check on the original streamed table would preserve an extra upstream enumeration. This buffer contains scalar cells, so the shallow behavior of [Table.Buffer](https://learn.microsoft.com/en-us/powerquery-m/table-buffer) is sufficient to materialize its values. Buffering may still increase memory or provide no speed benefit. It is a measured next experiment, not a claimed cure.

If `ExpandedParts` itself fails its bound, isolate `BufferedParts` first. If that scalar stage completes and the nested fallback expansion does not, the following replacement is a justified next candidate. If `BufferedParts` already fails, changing the later fallback construction is not supported by that result.

## Conditional replacement of lines 677–684

Retain existing code through `BufferedParts` at line 676. Replace only `ShiftPartJoin`, `PositionParts`, `ExpandedPositionParts` and `ExpandedParts` with the following named steps. Leave WorkSegments, WorkJoin and all segment business formulas unchanged. Do not combine this replacement with the additional buffer above in the first comparison; isolate the intervention.

```powerquery
    // A missing position in any part invalidates the whole shift's period allocation.
    PartPositionValidity = Table.Group(BufferedParts, {"SourceShiftID"}, {
        {"AllPositionsKnown", each
            List.NonNullCount([PartNetStartHours]) = Table.RowCount(_)
            and List.NonNullCount([PartNetEndHours]) = Table.RowCount(_), type logical}}),
    ValidShiftPositions = Table.SelectRows(PartPositionValidity, each [AllPositionsKnown]),
    ValidShiftKeys = Table.SelectColumns(ValidShiftPositions, {"SourceShiftID"}),
    ValidPartsJoin = Table.NestedJoin(BufferedParts, {"SourceShiftID"}, ValidShiftKeys, {"SourceShiftID"}, "ValidShift", JoinKind.Inner),
    ValidParts = Table.RemoveColumns(ValidPartsJoin, {"ValidShift", "EmployeeKey"}),
    MarkedValidParts = Table.AddColumn(ValidParts, "PartPositionsKnown", each true, type logical),
    ShiftPartJoin = Table.NestedJoin(ShiftColumns, {"SourceShiftID"}, MarkedValidParts, {"SourceShiftID"}, "Parts", JoinKind.LeftOuter),
    ExpandedValidParts = Table.ExpandTableColumn(ShiftPartJoin, "Parts",
        {"PayrollFortnightStart", "PartNetStartHours", "PartNetEndHours", "FortnightHoursBefore", "PartFortnightStatus", "PartPositionsKnown"},
        {"MatchedFortnightStart", "MatchedNetStartHours", "MatchedNetEndHours", "MatchedHoursBefore", "MatchedFortnightStatus", "PartPositionsKnown"}),
    PayrollFortnightStart = Table.AddColumn(ExpandedValidParts, "PayrollFortnightStart", each if [PartPositionsKnown] = true then [MatchedFortnightStart] else null),
    PartNetStartHours = Table.AddColumn(PayrollFortnightStart, "PartNetStartHours", each if [PartPositionsKnown] = true then [MatchedNetStartHours] else 0),
    PartNetEndHours = Table.AddColumn(PartNetStartHours, "PartNetEndHours", each if [PartPositionsKnown] = true then [MatchedNetEndHours] else [QualifyingNetHours]),
    FortnightHoursBefore = Table.AddColumn(PartNetEndHours, "FortnightHoursBefore", each if [PartPositionsKnown] = true then [MatchedHoursBefore] else null),
    PartFortnightStatus = Table.AddColumn(FortnightHoursBefore, "PartFortnightStatus", each if [PartPositionsKnown] = true then [MatchedFortnightStatus] else "Unknown"),
    ExpandedParts = Table.RemoveColumns(PartFortnightStatus, {"MatchedFortnightStart", "MatchedNetStartHours", "MatchedNetEndHours", "MatchedHoursBefore", "MatchedFortnightStatus", "PartPositionsKnown"}),
```

The current PositionParts column declares generic `type table`. The five proposed final scalar columns deliberately use default `any`, without imposing new coercions; exact intermediate type equality still needs the comparison below. Public ExcessSegments continues to impose its existing explicit thirteen-column output type. The marker distinguishes an unmatched row from a matched row whose non-position fields are null; valid null statuses or counters are preserved, not replaced by fallback values.

## Same-grain argument and limits

| Existing right-hand parts for a SourceShiftID | Current behavior | Proposed behavior |
|---|---|---|
| No rows | One fallback covering net positions 0 to QualifyingNetHours, with null period/counter and Unknown status | No valid key; left outer expansion retains one unmatched shift, then the same fallback values |
| One or more rows, all starts and ends non-null | Retain every original part and its five scalar fields | Unique valid key retains every original part exactly once; marker selects the original fields |
| Any row has a null start or end, including a mixture of valid and invalid parts | Discard all period allocation and produce one whole-shift fallback | The entire SourceShiftID fails validity, all its right-hand parts are removed, and left outer expansion yields one fallback |
| QualifyingNetHours is null and allocation falls back | One row with null net end; later cut construction yields no invented interval | Same null net end and unchanged cut construction |

The valid-key table is unique by grouping, so its inner join cannot duplicate a part. ShiftColumns remains unique by SourceShiftID. The following left outer join therefore produces either the original valid part count or one fallback per shift. [Microsoft's documented left outer merge-and-expand behavior](https://learn.microsoft.com/en-us/power-query/merge-queries-left-outer) retains unmatched left rows with null right-side values. Do not filter individual invalid parts and retain their siblings: that would violate the whole-shift fallback rule.

This removes the explicit per-shift `Table.SelectColumns` or `#table` provider created in PositionParts. A standard nested join still creates nested match tables; it does not eliminate all providers. Whether the engine executes this merge/expand pattern more efficiently is a runtime hypothesis. The extra validity grouping and inner join add work and are unwarranted if the smaller buffer experiment succeeds.

Before any approved implementation or sync, compare old/new ExpandedParts column names, types, row multiset and per-SourceShiftID counts; include no parts, all-valid split parts, mixed-validity parts, null hours and empty input. Then run the existing 57 native cases and a bounded complete ExcessSegments evaluation. Preserve all trigger, meal-boundary, timing, null/Unknown and interval-union assertions. Root owns implementation approval, native execution, synchronization and live refresh.
