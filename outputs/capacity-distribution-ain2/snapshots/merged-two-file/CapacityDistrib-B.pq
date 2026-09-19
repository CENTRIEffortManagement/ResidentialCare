section Section1;

// Query: fnAddIndexedRunningTotal
// Purpose: Calculate inclusive running totals in one pass through the existing indexed priority order.
// Inputs: Ordered rows with contiguous indexes starting at 2, a group-start index and a prefix count.
// Output: The original rows and columns followed by the requested total column, retaining type any.
// Notes: Buffer locally so validation, scalar projections and output rows use the same evaluated order.
shared fnAddIndexedRunningTotal = (
    orderedInput as table,
    valueColumn as text,
    indexColumn as text,
    startColumn as text,
    countColumn as text,
    outputColumn as text
) as table =>
let
    RequiredColumns = List.Distinct({valueColumn, indexColumn, startColumn, countColumn}),
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(orderedInput)),
    CheckedSchema =
        if not List.IsEmpty(MissingColumns) then
            error Error.Record("CapacityDistribB.RunningTotalSchema", "Required running-total columns are missing.", MissingColumns)
        else if Table.HasColumns(orderedInput, {outputColumn}) then
            error Error.Record("CapacityDistribB.RunningTotalSchema", "The running-total output column already exists.", outputColumn)
        else orderedInput,
    // These buffers apply to this invocation only; they do not cache separately refreshed queries.
    StableInput = Table.Buffer(CheckedSchema),
    Values = List.Buffer(Table.Column(StableInput, valueColumn)),
    Indexes = List.Buffer(Table.Column(StableInput, indexColumn)),
    GroupStarts = List.Buffer(Table.Column(StableInput, startColumn)),
    PrefixCounts = List.Buffer(Table.Column(StableInput, countColumn)),
    RowCount = Table.RowCount(StableInput),
    // Positional ranges require unique consecutive indexes and contiguous groups; never repair bad joins silently.
    IsValidControl = (position as number) as logical =>
        let
            RowIndex = Indexes{position},
            GroupStart = GroupStarts{position},
            PrefixCount = PrefixCounts{position}
        in
            try (
                Value.Is(RowIndex, type number) and RowIndex = position + 2
                and Value.Is(GroupStart, type number)
                and GroupStart >= 2 and GroupStart <= RowIndex and Number.RoundDown(GroupStart) = GroupStart
                and Value.Is(PrefixCount, type number) and PrefixCount = RowIndex - GroupStart + 1
                and (if position = 0 then GroupStart = RowIndex
                     else GroupStart = GroupStarts{position - 1} or GroupStart = RowIndex)
            ) otherwise false,
    FirstInvalidPosition = List.First(List.Select(List.Numbers(0, RowCount), each not IsValidControl(_)), null),
    ValidatedInput =
        if FirstInvalidPosition = null then StableInput
        else error Error.Record(
            "CapacityDistribB.RunningTotalOrder",
            "Running-total indexes or group boundaries are invalid. Check joins and the existing priority order.",
            [RowPosition = FirstInvalidPosition + 1, IndexColumn = indexColumn, StartColumn = startColumn, CountColumn = countColumn]
        ),
    // Emit one scalar per row without repeatedly slicing/summing prefixes or appending to a growing list.
    // List.Sum of the two scalars preserves null for an entirely null prefix and ignores null otherwise.
    RunningTotals = List.Buffer(List.Generate(
        () => [Position = 0, Total = if RowCount = 0 then null else Values{0}],
        (state as record) as logical => state[Position] < RowCount,
        (state as record) as record =>
            let NextPosition = state[Position] + 1
            in [
                Position = NextPosition,
                Total = if NextPosition >= RowCount then null
                        else if GroupStarts{NextPosition} <> GroupStarts{state[Position]} then Values{NextPosition}
                        else List.Sum({state[Total], Values{NextPosition}})
            ],
        (state as record) => state[Total]
    )),
    WithRunningTotal = Table.AddColumn(
        ValidatedInput, outputColumn,
        (row as record) => RunningTotals{Record.Field(row, indexColumn) - 2},
        type any
    )
in
    WithRunningTotal;


shared #"ResMaxReduction (C# Unassigned)" = let
    Source = ResPeriodOverallocationReduction,
    #"Grouped RESMAXADDITION" = Table.Group(Source, {"Resource"}, {{"MaxResAddition", each List.Sum([Unassigned]), type nullable number}}),
    #"Merged Queries1" = Table.NestedJoin(#"Grouped RESMAXADDITION", {"Resource"}, ResRosterAvailabilityCapReduction, {"Resource"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction" = Table.ExpandTableColumn(#"Merged Queries1", "ResRosterAvailabilityCapReduction", {"AvailabilityReduction"}, {"AvailabilityReduction"}),
    #"Inserted Minimum1" = Table.AddColumn(#"Expanded ResRosterAvailabilityCapReduction", "ResMaxReduction", each List.Min({[MaxResAddition], [AvailabilityReduction]}), type number),
    #"Sorted Rows" = Table.Sort(#"Inserted Minimum1",{{"Resource", Order.Ascending}})
in
    #"Sorted Rows";

shared ResPeriodMaxAddition = let
    Source = Table.NestedJoin(ResPeriodOverallocationReduction, {"Resource"}, #"ResMaxReduction (C# Unassigned)", {"Resource"}, "ResMaxAddition", JoinKind.LeftOuter),
    #"Expanded ResMaxAddition" = Table.ExpandTableColumn(Source, "ResMaxAddition", {"ResMaxReduction"}, {"ResMaxReduction"})
in
    #"Expanded ResMaxAddition";

[ Description = "%#(lf)(A-D) gaps divided to D #(lf)#(lf)Not really ABi - but should still provide same ranking of prorrity" ]
shared PeriodUnderallcationPrioritised = let
    Source = ResPeriodOverallocationReduction,
    #"Merged Queries2" = Table.NestedJoin(Source, {"Resource", "Period"}, ResPeriodMaxAddition, {"Resource", "Period"}, "ResPeriodMaxAddition", JoinKind.LeftOuter),
    #"Expanded ResPeriodMaxAddition" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodMaxAddition", {"ResMaxReduction"}, {"ResMaxReduction"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResPeriodMaxAddition", {"Period"}, PeriodDemandTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded PeriodDemandTABLE", {"Resource", "Period"}, #"IMPORT ResPeriodNWDTABLE", {"Resource", "Period"}, "IMPORT ResPeriodNWDTABLE", JoinKind.LeftOuter),
    #"Expanded NSWPRIORITIES" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ResPeriodNWDTABLE", {"NWDPriority"}, {"NWDPriority"}),
    BUFFER = Table.Buffer(#"Expanded NSWPRIORITIES")
in
    BUFFER;

shared #"PrioritiseRedistribAvail-Setup" = let
    Source = #"PeriodUnderallcationPrioritised",
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 2, 1, Int64.Type),
    #"Added FTE" = Table.AddColumn(#"Added Index", "FTE", each 1)
in
    #"Added FTE";

shared #"PrioritiseRedistribAvail-Distrib" = let
    Source = #"PrioritiseRedistribAvail-Setup",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Index", "IndexAllRows"}}),
    // Period is the stable tie-breaker inside an otherwise-equal Resource priority.
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"Resource", Order.Ascending}, {"C#/D", Order.Descending}, {"ExcessC#-D", Order.Descending}, {"NWDPriority", Order.Descending}, {"Period", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRowPrioritySort", 2
, 1, Int64.Type)
in
    #"Added Index";

shared #"ResIndexLimits (ResMaxReduction)" = let
    Source = #"PrioritiseRedistribAvail-Distrib",
    #"Added Index" = Table.AddIndexColumn(Source, "Index2", 1, 1, Int64.Type),
    #"Grouped Rows" = Table.Group(#"Added Index", {"Resource"}, {{"StartResIndex", each List.Min([IndexAllRowPrioritySort]), type number}, {"ResExcessAvailable", each List.Max([ResMaxReduction]), type nullable number}, {"ResAvailNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

// Query: ReDistributeResAvailability
// Purpose: Apply each resource's reduction cap in the existing redistribution priority order.
shared ReDistributeResAvailability = let
    Source = Table.NestedJoin(#"PrioritiseRedistribAvail-Distrib", {"Resource"}, #"ResIndexLimits (ResMaxReduction)", {"Resource"}, "REsLimits", JoinKind.LeftOuter),
    #"Expanded REsLimits" = Table.ExpandTableColumn(Source, "REsLimits", {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}, {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded REsLimits", {"Resource", "Period"}, #"ResPeriodAvailabilityCapped(C#)TABLE", {"Resource", "Period"}, "ResPeriodAvailabilityCapped(C#)TABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAvailabilityCapped(C#)TABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAvailabilityCapped(C#)TABLE", {"C#"}, {"C#"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded ResPeriodAvailabilityCapped(C#)TABLE",{{"IndexAllRowPrioritySort", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows1", "Subtraction", each [IndexAllRowPrioritySort] - [StartResIndex]+1),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Subtraction",{{"Subtraction", Int64.Type}}),
    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Changed Type", "C#", "IndexAllRowPrioritySort", "StartResIndex", "Subtraction", "RunningResTotal"),
    #"Added KEEP" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "KeepRunningTotal", each if [RunningResTotal] > [ResMaxReduction] then "Remove" else if [RunningResTotal] <= [ResMaxReduction] then "Keep" else "Split")
in
    #"Added KEEP";

[ Description = "BUFFER" ]
shared ReDistribResAvailabilityTABLE = let
    Source = ReDistributeResAvailability,
    #"Sorted Rows" = Table.Sort(Source,{{"Period", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Index", 2, 1, Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Index",{"D", "NWDPriority", "IndexAllRows", "FTE", "IndexAllRowPrioritySort", "StartResIndex", "ResExcessAvailable", "ResAvailNumber", "Subtraction"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER
;

shared #"PeriodIndexLimits (ExcessC#-D)" = let
    Source = ReDistribResAvailabilityTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([Index]), type number}, {"MaxC#-D", each List.Max([#"ExcessC#-D"]), type nullable number}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
// Query: ReDistribPeriodAvailbility TABLE
// Purpose: Retain redistribution rows within both resource and period reduction limits.
shared #"ReDistribPeriodAvailbility TABLE" = let
    Source = Table.NestedJoin(ReDistribResAvailabilityTABLE, {"Period"}, #"PeriodIndexLimits (ExcessC#-D)", {"Period"}, "PeriodIndexLimits", JoinKind.LeftOuter),
    #"Expanded PeriodIndexLimits" = Table.ExpandTableColumn(Source, "PeriodIndexLimits", {"StartPeriodIndex"}, {"StartPeriodIndex"}),
    #"Added Custom" = Table.AddColumn(#"Expanded PeriodIndexLimits", "ApplicableC#", each if [KeepRunningTotal] = "Remove" then 0 else [#"C#"]),
    #"Sorted INDEX" = Table.Sort(#"Added Custom",{{"Index", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted INDEX", "Subtraction", each [Index] - [StartPeriodIndex]+1, type number),
    #"Added PERIODRUNNINGTOTAL" = fnAddIndexedRunningTotal(#"Inserted Subtraction", "ApplicableC#", "Index", "StartPeriodIndex", "Subtraction", "PeriodRunningTotal"),
    #"Added KEEP PR" = Table.AddColumn(#"Added PERIODRUNNINGTOTAL", "Keep RunningTotal", each if [RunningResTotal] <= [ResMaxReduction]
and 
[PeriodRunningTotal] <=[#"ExcessC#-D"]
then "Keep" else "Remove"),
    #"Filtered Rows" = Table.SelectRows(#"Added KEEP PR", each ([Keep RunningTotal] = "Keep")),
    BUFFER = Table.Buffer(#"Filtered Rows")
in
    BUFFER;

shared CheckReAllocation = let
    Source = #"ReDistribPeriodAvailbility TABLE",
    #"Grouped Rows1" = Table.Group(Source, {"Resource"}, {{"Max", each List.Max([RunningResTotal]), type number}, {"ReDirect", each List.Sum([#"C#"]), type nullable number}})
in
    #"Grouped Rows1";

shared ReAllocateMATRIX = let
    Source = ReDistribResAvailabilityTABLE,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Period", "C#"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Period", type text}}),
    #"Pivoted Column1" = Table.Pivot(#"Changed Type1", List.Distinct(#"Changed Type1"[Period]), "Period", "C#", List.Sum)
in
    #"Pivoted Column1";

shared #"ResRosterAvailabilityC##CapReduction" = let
    Source = #"C##TABLE",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"RosterAvailability", each List.Sum([#"C##"]), type nullable number}}),
    #"Replaced Value" = Table.ReplaceValue(#"Grouped Rows",null,0,Replacer.ReplaceValue,{"RosterAvailability"}),
    // Join the cap only after aggregation so contract totals are never repeated at Resource-Period grain.
    #"Merged Resource Contract" = Table.NestedJoin(#"Replaced Value", {"Resource"}, BResourceContract, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    #"Expanded Effective Shift Cap" = Table.ExpandTableColumn(#"Merged Resource Contract", "ResourceContract", {"Effective Shift Cap"}, {"Effective Shift Cap"}),
    #"Added AVAILCAP" = Table.AddColumn(#"Expanded Effective Shift Cap", "RosterAvailabilityCAPPED", each
        // Allocation-only Resources can remain in the grid with zero availability and no contract cap.
        if [RosterAvailability] = 0 then 0
        else if [Effective Shift Cap] = null then error Error.Record(
            "CapacityDistribB.MissingResourceCap",
            "A Resource with positive C## availability has no Effective Shift Cap.",
            [Resource = [Resource], RosterAvailability = [RosterAvailability]]
        )
        else if [RosterAvailability] > [Effective Shift Cap] then [Effective Shift Cap]
        else [RosterAvailability], type number),
    #"Inserted Subtraction" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"RosterAvailability", "RosterAvailabilityCAPPED", "Effective Shift Cap"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailabilityReduction", "MaxC##Reduction"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([#"MaxC##Reduction"] <> 0))
in
    #"Filtered Rows";

// Query: RoleFolder
// Purpose: Expose the physical calculation-folder name resolved from this workbook's own FilePathUrl.
shared RoleFolder = let
    MatchingRows = Table.SelectRows(RolePathTABLE, each [Variable Name] = "Role"),
    Value = if Table.RowCount(MatchingRows) = 1 then MatchingRows{0}[Value]
        else error "RolePathTABLE must contain exactly one physical Role row."
in
    Value;

// Query: SourceRole
// Purpose: Keep AIN as the upstream business role while the isolated workbook physically runs in AIN2.
shared SourceRole =
    if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN") = 0 or Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 then "AIN"
    else error Error.Record("CapacityDistribB.UnsupportedRoleFolder", "This AIN build must run only from physical folder AIN or AIN2.", RoleFolder);

// Query: OutputRole
// Purpose: Publish AIN2 labels during development and AIN labels automatically after the approved folder rename.
shared OutputRole =
    if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 then "AIN2"
    else if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN") = 0 then "AIN"
    else error Error.Record("CapacityDistribB.UnsupportedRoleFolder", "This AIN build must run only from physical folder AIN or AIN2.", RoleFolder);

// Query: Role
// Purpose: Preserve the existing compatibility scalar while publishing the explicit output role.
shared Role = OutputRole;

// Query: RoleContext_CHECK
// Purpose: Validate the only two supported physical/source/output tuples for this AIN build.
// Output: Connection-only check rows forced by the final B publication gate.
shared RoleContext_CHECK = let
    ExpectedSourceRole = "AIN",
    ExpectedOutputRole = if Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 then "AIN2" else "AIN",
    SupportedFolder = Comparer.OrdinalIgnoreCase(RoleFolder, "AIN2") = 0 or Comparer.OrdinalIgnoreCase(RoleFolder, "AIN") = 0,
    SourceAttempt = try SourceRole,
    OutputAttempt = try OutputRole,
    ValidTuple = SupportedFolder
        and not SourceAttempt[HasError]
        and not OutputAttempt[HasError]
        and SourceAttempt[Value] = ExpectedSourceRole
        and OutputAttempt[Value] = ExpectedOutputRole,
    Checks = #table(
        type table [Check = text, Status = text, Actual = nullable text, Expected = nullable text],
        {
            {"Supported RoleFolder", if SupportedFolder then "Pass" else "Fail", RoleFolder, "AIN2 or AIN"},
            {"SourceRole", if SourceRole = ExpectedSourceRole then "Pass" else "Fail", SourceRole, ExpectedSourceRole},
            {"OutputRole", if OutputRole = ExpectedOutputRole then "Pass" else "Fail", OutputRole, ExpectedOutputRole}
        }
    ),
    Checked = if ValidTuple then Checks
        else error Error.Record(
            "CapacityDistribB.RoleContext",
            "The physical/source/output role tuple is unsupported.",
            [RoleFolder = RoleFolder, ExpectedSourceRole = ExpectedSourceRole, ExpectedOutputRole = ExpectedOutputRole]
        )
in
    Checked;

// Query: fnIntraLaneWorkbookPath
// Purpose: Build workbook paths only inside the physical role folder, never by substituting SourceRole.
shared fnIntraLaneWorkbookPath = (workbookName as text) as text =>
let
    SafeName = if workbookName = Text.AfterDelimiter(workbookName, "\", {0, RelativePosition.FromEnd})
        and not Text.Contains(workbookName, "/")
        and workbookName <> ""
        then workbookName
        else error Error.Record("CapacityDistribB.InvalidIntraLaneFile", "An intra-lane import must be a file name without a path.", workbookName),
    Candidate = RolePath & "\" & SafeName,
    CandidateFolder = Text.BeforeDelimiter(Candidate, "\", {0, RelativePosition.FromEnd}),
    Checked = if Comparer.OrdinalIgnoreCase(CandidateFolder, RolePath) = 0 then Candidate
        else error Error.Record("CapacityDistribB.IntraLaneEscape", "An intra-lane import resolved outside the physical role folder.", Candidate)
in
    Checked;

// Query: IMPORTSource Settings Data
// Purpose: Import the shared Settings Data workbook for this unit.
shared #"IMPORTSource Settings Data" = let
    // Settings Data is in the parent Calculations folder for every role.
    CalculationsPath = Text.BeforeDelimiter(
        RolePath, "\", {0, RelativePosition.FromEnd}
    ),
    Source = Excel.Workbook(
        File.Contents(CalculationsPath & "\Settings Data.xlsx"),
        null,
        true
    )
in
    Source;

// Query: EXTRACT MaxAvailability
// Purpose: Extract the analysis-period worker shift cap from Settings Data.
shared #"EXTRACT MaxAvailability" = let
    SettingsTable = #"IMPORTSource Settings Data"{
        [Item = "MaxAvailability", Kind = "Table"]
    }[Data],
    TypedSettings = Table.TransformColumnTypes(
        SettingsTable, {{"MaxAvailability", Int64.Type}}
    ),
    PeriodShiftLimit = TypedSettings{0}[MaxAvailability]
in
    PeriodShiftLimit;

// Query: MaxAvailability
// Purpose: Preserve the existing cap interface for downstream B calculations.
shared MaxAvailability = #"EXTRACT MaxAvailability";

shared AllocationThreshold = 0.8 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

// Query: BResourceContract
// Purpose: Publish one validated Effective Shift Cap per Resource for B's Resource-grain calculations.
shared BResourceContract = let
    Failures = Table.SelectRows(BResourceContract_CHECK, each [Status] <> "Pass"),
    Checked = if Table.IsEmpty(Failures) then #"IMPORT ResourceContract"
        else error Error.Record(
            "CapacityDistribB.ResourceContractValidation",
            "Resource contracts are missing, duplicated or invalid. Review BResourceContract_CHECK.",
            Failures
        ),
    Output = Table.SelectColumns(Checked, {"Resource", "Effective Shift Cap", "Limit Basis"})
in
    Table.Buffer(Output);

// Query: IMPORT Demand
// Purpose: Read the existing A.1 Demand_Prepare table without changing its schema or grain.
shared #"IMPORT Demand" = let
    Source = #"IMPORTSource A1",
    Demand_Prepare_Table = Source{[Item="Demand_Prepare",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(Demand_Prepare_Table, {"D", "Period"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"D", Int64.Type}, {"Period", Int64.Type}})
in
    #"Changed Type";

// Query: IMPORT Allocation
// Purpose: Read A.1 resource-period allocations, preserving the existing null-period exclusion.
shared #"IMPORT Allocation" = let
    Source = #"IMPORTSource A1",
    ResPeriodAllocationTABLE_Table = Source{[Item="ResPeriodAllocationTABLE",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResPeriodAllocationTABLE_Table, {"Allocation", "Resource", "Period"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Allocation", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Period] <> null))
in
    #"Filtered Rows";

[ Description = "BUFFER" ]
// Query: IMPORT AvailabilityOriginal
// Purpose: Read and retain the existing buffered original-availability table from A.1.
shared #"IMPORT AvailabilityOriginal" = let
    Source = #"IMPORTSource A1",
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResPeriodAvailabilityTABLE_Table, {"Role", "Resource", "Period", "Availability"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    BUFFER = Table.Buffer(#"Changed Type")
in
    BUFFER;

// Query: IMPORT ResPeriodAvailabilityTABLE
// Purpose: Reuse B's typed original-availability extraction as the merged A2 cap input.
shared #"IMPORT ResPeriodAvailabilityTABLE" = #"IMPORT AvailabilityOriginal";

// Query: IMPORT ResPeriodCapPrioritised
// Purpose: Read the narrow A1 period-cap priority interface used by the merged A2 stages.
shared #"IMPORT ResPeriodCapPrioritised" = let
    Source = #"IMPORTSource A1",
    ResPeriodCapPrioritised_Table = Source{[Item="ResPeriodCapPrioritised",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResPeriodCapPrioritised_Table, {"Resource", "Period", "Allocation", "PotentialAvailability", "PRCell", "UnassignedAvail", "ResAv-C'", "PeriodA-CNeg", "C'-D", "C'/D", "ClusterAllocation", "NWDPriority", "NWDType"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"Allocation", type any}, {"PotentialAvailability", type any}, {"PRCell", type text}, {"UnassignedAvail", Int64.Type}, {"ResAv-C'", Int64.Type}, {"PeriodA-CNeg", type number}, {"C'-D", Int64.Type}, {"C'/D", type number}, {"ClusterAllocation", type number}, {"NWDPriority", Int64.Type}, {"NWDType", type text}})
in
    #"Changed Type";

// Query: IMPORT MultiPeriod - Remove
// Purpose: Read only the A1 multi-period keep flag used by the merged capped-availability rule.
shared #"IMPORT MultiPeriod - Remove" = let
    Source = #"IMPORTSource A1",
    MultiDayPeriod_Remove_Table = Source{[Item="MultiDayPeriod_Remove",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(MultiDayPeriod_Remove_Table, {"Resource", "AvailablePeriod", "NoAllocationKeep"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"NoAllocationKeep", type logical}})
in
    #"Changed Type";

// Query: IMPORT ResPeriodWDTABLE
// Purpose: Read only roster status and availability at Resource x Period grain from A1.
shared #"IMPORT ResPeriodWDTABLE" = let
    Source = #"IMPORTSource A1",
    ResPeriodWDTABLE_Table = Source{[Item="ResPeriodWDTABLE",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResPeriodWDTABLE_Table, {"Resource", "Period", "Availability", "RosteredPeriodStatus"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"RosteredPeriodStatus", type text}})
in
    #"Changed Type";

// Query: IMPORT ResourcePeriodTABLE_empty
// Purpose: Read the authoritative narrow Role x Resource x Period skeleton from A1.
shared #"IMPORT ResourcePeriodTABLE_empty" = let
    Source = #"IMPORTSource A1",
    ResourcePeriodTABLE_empty_Table = Source{[Item="ResourcePeriodTABLE_empty",Kind="Table"]}[Data],
    Selected = Table.SelectColumns(ResourcePeriodTABLE_empty_Table, {"Role", "Resource", "Period"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}})
in
    #"Changed Type";

// Query: IMPORT ResPeriodNWDTABLE
// Purpose: Read the existing A.1 resource-period priority worksheet, preserving header promotion and types.
shared #"IMPORT ResPeriodNWDTABLE" = let
    Source = #"IMPORTSource A1",
    ResPeriodWDTABLE_Sheet = Source{[Item="ResPeriodWDTABLE",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(ResPeriodWDTABLE_Sheet, [PromoteAllScalars=true]),
    Selected = Table.SelectColumns(#"Promoted Headers", {"Resource", "Period", "NWDPriority"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"NWDPriority", Int64.Type}})
in
    #"Changed Type";

// Query: IMPORTSource StaffListMaster
// Purpose: Import the Unit1 StaffListMaster workbook once for B's Resource-grain contract caps.
shared #"IMPORTSource StaffListMaster" = let
    CalculationsPath = Text.BeforeDelimiter(RolePath, "\", {0, RelativePosition.FromEnd}),
    WorkbookBinary = Binary.Buffer(File.Contents(CalculationsPath & "\StaffListMaster.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in
    Table.Buffer(WorkbookNavigation);

// Query: IMPORTSource A1
// Purpose: Provide the A.1 workbook binary and navigation table for the four existing imports.
// Notes: Navigation buffering is shallow; selected table transformations remain in their owning imports.
shared #"IMPORTSource A1" = let
    // Reuse the saved file bytes within an evaluation; separately loaded outputs can still evaluate this again.
    WorkbookBinary = Binary.Buffer(File.Contents(fnIntraLaneWorkbookPath("CapacityDistrib(A.1)-shifts.xlsx"))),
    WorkbookNavigation = Table.Buffer(Excel.Workbook(WorkbookBinary, null, true))
in
    WorkbookNavigation;

// Query: IMPORT ResourceContract
// Purpose: Extract the validated preferred-role contract fields from Table_Masterlist for this role.
// Output: One candidate contract row per preferred AIN Resource before B's independent checks.
shared #"IMPORT ResourceContract" = let
    Matches = Table.SelectRows(#"IMPORTSource StaffListMaster", each [Item] = "Table_Masterlist" and [Kind] = "Table"),
    MasterlistTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("B resource contract import", "Expected exactly one Table_Masterlist table in StaffListMaster.xlsx.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Resource", "Role", "PreferredRole", "Effective Shift Cap", "Limit Basis"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(MasterlistTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(MasterlistTable, RequiredColumns)
        else error Error.Record("B resource contract import", "Table_Masterlist is missing required contract columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected, {
        {"Resource", Int64.Type}, {"Role", type text}, {"PreferredRole", type text},
        {"Effective Shift Cap", type number}, {"Limit Basis", type text}
    }),
    NormalizedRoles = Table.TransformColumns(Typed, {
        {"Role", each if _ = null then null else Text.Clean(Text.Trim(_)), type nullable text},
        {"PreferredRole", each if _ = null then null else Text.Clean(Text.Trim(_)), type nullable text}
    }),
    RoleValue = SourceRole,
    // Preserve Worker Reconciliation's preferred-role decision; no contract is split across roles.
    PreferredRoleContracts = Table.SelectRows(NormalizedRoles, each [Role] = RoleValue and [PreferredRole] = RoleValue)
in
    Table.Buffer(PreferredRoleContracts);

// Query: BResourceContract_CHECK
// Purpose: Validate unique Resource contracts, usable caps and complete coverage of AIN availability resources.
// Output: Connection-only diagnostic rows used by BResourceContract and the final B publication gate.
shared BResourceContract_CHECK = let
    Contracts = #"IMPORT ResourceContract",
    ContractKeys = Table.SelectColumns(Contracts, {"Resource"}),
    ErrorKeyRows = Table.RowCount(Table.SelectRowsWithErrors(ContractKeys, {"Resource"})),
    KeysWithoutErrors = Table.RemoveRowsWithErrors(ContractKeys, {"Resource"}),
    MissingKeyRows = Table.RowCount(Table.SelectRows(KeysWithoutErrors, each [Resource] = null)),
    CompleteKeys = Table.SelectRows(KeysWithoutErrors, each [Resource] <> null),
    ContractCounts = Table.Group(CompleteKeys, {"Resource"}, {{"ContractRows", each Table.RowCount(_), Int64.Type}}),
    DuplicateResourceGroups = Table.RowCount(Table.SelectRows(ContractCounts, each [ContractRows] > 1)),
    ContractValues = Table.SelectColumns(Contracts, {"Effective Shift Cap", "Limit Basis"}),
    ErrorContractRows = Table.RowCount(Table.SelectRowsWithErrors(ContractValues, {"Effective Shift Cap", "Limit Basis"})),
    ContractValuesWithoutErrors = Table.RemoveRowsWithErrors(ContractValues, {"Effective Shift Cap", "Limit Basis"}),
    InvalidCapRows = ErrorContractRows + Table.RowCount(Table.SelectRows(ContractValuesWithoutErrors, each
        [Effective Shift Cap] = null
            or [Effective Shift Cap] <= 0
            or [Effective Shift Cap] <> Number.RoundDown([Effective Shift Cap])
            or [Limit Basis] = null
            or Text.Trim([Limit Basis]) = "")),
    // Coverage is assessed at Resource grain before any Resource-Period distribution joins.
    AvailabilityResources = Table.Distinct(Table.SelectColumns(
        // Only Resources contributing positive availability require contract coverage.
        Table.SelectRows(#"IMPORT AvailabilityOriginal", each
            [Resource] <> null and [Availability] <> null and [Availability] > 0),
        {"Resource"}
    )),
    MissingCoverage = Table.RowCount(Table.NestedJoin(
        AvailabilityResources, {"Resource"}, CompleteKeys, {"Resource"}, "Contract", JoinKind.LeftAnti
    )),
    Checks = #table(
        type table [Check = text, Status = text, Failures = number, Details = nullable text],
        {
            {"Populated Resource contract keys", if ErrorKeyRows + MissingKeyRows = 0 then "Pass" else "Fail", ErrorKeyRows + MissingKeyRows, null},
            {"Unique Resource contracts", if DuplicateResourceGroups = 0 then "Pass" else "Fail", DuplicateResourceGroups, null},
            {"Valid effective Resource caps", if InvalidCapRows = 0 then "Pass" else "Fail", InvalidCapRows, null},
            {"Complete AIN availability contract coverage", if MissingCoverage = 0 then "Pass" else "Fail", MissingCoverage, null}
        }
    )
in
    Table.Buffer(Checks);

shared PeriodAllocationTABLE = let
    Source = ResPeriodAllocationTABLE,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Period", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Allocation", "AllocationX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "Allocation", each if [AllocationX] > AllocationThreshold then 1 else [AllocationX]),
    #"Grouped Rows" = Table.Group(#"Added Conditional Column", {"Period"}, {{"A", each List.Sum([Allocation]), type number}})
in
    #"Grouped Rows";

shared #"PeriodA-D" = let
    Source = Table.NestedJoin(PeriodDemandTABLE, {"Period"}, PeriodAllocationTABLE, {"Period"}, "AllocationPeriodLIST", JoinKind.LeftOuter),
    #"Expanded AllocationPeriodLIST" = Table.ExpandTableColumn(Source, "AllocationPeriodLIST", {"A"}, {"A"}),
    #"Sorted Rows" = Table.Sort(#"Expanded AllocationPeriodLIST",{{"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared ResRosterAvailabilityCapReduction = let
    Source = ResPeriodAvailabilityTABLE,
    #"Grouped ROSTERAVAILABILITY" = Table.Group(Source, {"Resource"}, {{"RosterAvailability", each List.Sum([#"C#"]), type nullable number}}),
    #"Replaced Value" = Table.ReplaceValue(#"Grouped ROSTERAVAILABILITY",null,0,Replacer.ReplaceValue,{"RosterAvailability"}),
    // Join the cap only after aggregation so contract totals are never repeated at Resource-Period grain.
    #"Merged Resource Contract" = Table.NestedJoin(#"Replaced Value", {"Resource"}, BResourceContract, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    #"Expanded Effective Shift Cap" = Table.ExpandTableColumn(#"Merged Resource Contract", "ResourceContract", {"Effective Shift Cap"}, {"Effective Shift Cap"}),
    #"Added AVAILCAP" = Table.AddColumn(#"Expanded Effective Shift Cap", "RosterAvailabilityCAPPED", each
        // Allocation-only Resources can remain in the grid with zero availability and no contract cap.
        if [RosterAvailability] = 0 then 0
        else if [Effective Shift Cap] = null then error Error.Record(
            "CapacityDistribB.MissingResourceCap",
            "A Resource with positive C# availability has no Effective Shift Cap.",
            [Resource = [Resource], RosterAvailability = [RosterAvailability]]
        )
        else if [RosterAvailability] > [Effective Shift Cap] then [Effective Shift Cap]
        else [RosterAvailability], type number),
    #"Inserted Subtraction" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"RosterAvailability", "RosterAvailabilityCAPPED", "Effective Shift Cap"})
in
    #"Removed Columns";

[ Description = "BUFFER" ]
shared ResPeriodOverallocationReduction = let
    // Period A-D Overallocated
    Source = #"PeriodA-DPos (Overallocation)",
    BUFFER = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Period"}, PeriodCapacityTABLE, {"Period"}, "PeriodCapacityTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodCapacityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodCapacityTABLE", {"PeriodCapacity"}, {"PeriodCapacity"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodCapacityTABLE", {"Period"}, #"PeriodC#-DPos (Excess)TABLE", {"Period"}, "PeriodC#-DPos (Excess)TABLE", JoinKind.LeftOuter),
    #"Expanded PeriodC#-DPos (Excess)TABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodC#-DPos (Excess)TABLE", {"ExcessC#-D", "C#/D"}, {"ExcessC#-D", "C#/D"}),
    #"Merged Queries3" = Table.NestedJoin(#"Expanded PeriodC#-DPos (Excess)TABLE", {"Period"}, #"ResourcePeriods-EmptyTABLE", {"Period"}, "ResourcePeriods-EmptyTABLE", JoinKind.LeftOuter),
    #"Expanded ResourcePeriods-EmptyTABLE" = Table.ExpandTableColumn(#"Merged Queries3", "ResourcePeriods-EmptyTABLE", {"Resource"}, {"Resource"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Expanded ResourcePeriods-EmptyTABLE",{"Resource", "Period", "D", "A", "A-D", "PeriodCapacity", "ExcessC#-D", "C#/D"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns1",{{"Resource", Int64.Type}}),
    #"Merged Queries2" = Table.NestedJoin(#"Changed Type", {"Period", "Resource"}, ResPeriodAllocationTABLE, {"Period", "Resource"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Merged Queries5" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource"}, ResRosterAvailabilityCapReduction, {"Resource"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction1" = Table.ExpandTableColumn(#"Merged Queries5", "ResRosterAvailabilityCapReduction", {"AvailabilityReduction"}, {"ResAvailabilityReduction"}),
    #"Merged Queries4" = Table.NestedJoin(#"Expanded ResRosterAvailabilityCapReduction1", {"Resource", "Period"}, ResPeriodAvailabilityTABLE, {"Resource", "Period"}, "ResRosterAvailabilityCapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityCapReduction" = Table.ExpandTableColumn(#"Merged Queries4", "ResRosterAvailabilityCapReduction", {"C#"}, {"C#"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResRosterAvailabilityCapReduction",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    /*#"Filtered Rows1" = Table.SelectRows(#"Expanded ResRosterAvailabilityCapReduction", each ([#"C#"] = 1) and ([ResAvailabilityReduction] <> null)),
   */ 
   #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each 
        ([ResAvailabilityReduction] <> null 
        and [ResAvailabilityReduction] <> 0) 
        and ([#"C#/D"] <> null 
        and [#"C#/D"] <> 0) 
        and ([#"C#"] <> 0) 
        and ([#"C#"] <> null
        and (([Allocation] = null )
             or ([Allocation] <> 0 ))  )
    ),
    #"Renamed Columns1" = Table.RenameColumns(#"Filtered Rows",{{"C#", "Unassigned"}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Renamed Columns1",{"Resource", "Period", "ExcessC#-D", "C#/D", "Unassigned"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Other Columns",{"Resource", "Period", "Unassigned"})
in
    #"Reordered Columns";

shared ResPeriodAllocationTABLE = let
    Source = #"IMPORT Allocation",
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Resource", Int64.Type}})
in
    #"Changed Type1";

shared AvailabilityOriginalMATRIX = let
    Source = #"ResourcePeriods-EmptyTABLE",
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, #"IMPORT AvailabilityOriginal", {"Resource", "Period"}, "IMPORT AvailabilityOriginal", JoinKind.LeftOuter),
    #"Expanded IMPORT AvailabilityOriginal" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT AvailabilityOriginal", {"Availability"}, {"Availability"}),
    #"Sorted Rows" = Table.Sort(#"Expanded IMPORT AvailabilityOriginal",{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "Availability", List.Sum)
in
    #"Pivoted Column";

shared PeriodDemandTABLE = let
    Source = #"IMPORT Demand",
    // Downstream B calculations need only one Demand value at Period grain.
    Selected = Table.SelectColumns(Source, {"Period", "D"}),
    #"Changed Type" = Table.TransformColumnTypes(Selected,{{"Period", Int64.Type}})
in
    #"Changed Type";

shared PeriodDemandMATRIX = let
    Source = PeriodDemandTABLE,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Period", "D"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Other Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Other Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "D", List.Sum)
in
    #"Pivoted Column";

// Query: PeriodCapPrioritisedTABLE
// Purpose: Establish deterministic per-period removal priority for the merged A2 cap stage.
shared PeriodCapPrioritisedTABLE = let
    Source = #"IMPORT ResPeriodCapPrioritised",
    // Resource and PRCell break otherwise-equal priority rows without changing the business priority fields.
    #"Sorted NWD,C-D,PA-Cn,PERIOD" = Table.Sort(Source,{{"Period", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"ClusterAllocation", Order.Descending}, {"C'-D", Order.Ascending}, {"PeriodA-CNeg", Order.Ascending}, {"Resource", Order.Ascending}, {"PRCell", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted NWD,C-D,PA-Cn,PERIOD", "PeriodIndex", 2, 1, Int64.Type)
in
    #"Added Index";

// Query: PeriodCapIndexLimits
// Purpose: Capture each period's 2-based start index for the linear running total.
shared PeriodCapIndexLimits = let
    Source = PeriodCapPrioritisedTABLE,
    #"Grouped INDEX+COUNT" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([PeriodIndex]), Int64.Type}})
in
    #"Grouped INDEX+COUNT";

// Query: Period Running Total
// Purpose: Apply the period cap with one linear pass over the deterministic priority order.
shared #"Period Running Total" = let
    Source = Table.NestedJoin(PeriodCapPrioritisedTABLE, {"Period"}, PeriodCapIndexLimits, {"Period"}, "PeriodCapLimits", JoinKind.LeftOuter),
    #"Expanded PeriodLimits" = Table.ExpandTableColumn(Source, "PeriodCapLimits", {"StartPeriodIndex"}, {"StartPeriodIndex"}),
    #"Sorted Rows" = Table.Sort(#"Expanded PeriodLimits",{{"PeriodIndex", Order.Ascending}}),
    // The removable amount is limited by both the period availability gap and demand excess.
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Sorted Rows", "MaxPeriodReduction", each List.Min({-1 * [#"PeriodA-CNeg"], [#"C'-D"]})),
    #"Added Prefix Count" = Table.AddColumn(#"Inserted MAXRESTOTAL", "PeriodPrefixCount", each [PeriodIndex] - [StartPeriodIndex] + 1, Int64.Type),
    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Added Prefix Count", "UnassignedAvail", "PeriodIndex", "StartPeriodIndex", "PeriodPrefixCount", "PeriodRunningTotal"),
    #"Renamed Columns" = Table.RenameColumns(#"Added RUNNINGRESTOTAL",{{"UnassignedAvail", "Reduction"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "ReducePeriod", each if [PeriodRunningTotal] <= [MaxPeriodReduction] then "RemovableByPeriod" else "UnremovableByPeriod")
in
    #"Added Conditional Column";

// Query: ResCapPrioritisedTABLE
// Purpose: Establish deterministic per-resource removal priority for the merged A2 cap stage.
shared ResCapPrioritisedTABLE = let
    Source = #"Period Running Total",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Period", "PRCell", "Reduction", "ResAv-C'", "ClusterAllocation", "NWDPriority", "NWDType", "MaxPeriodReduction", "PeriodRunningTotal", "ReducePeriod"}),
    // Period and PRCell break otherwise-equal priority rows without changing the business priority fields.
    #"Sorted RES,NWDPRI,MAXPRT,PRT" = Table.Sort(#"Removed Other Columns",{{"Resource", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"MaxPeriodReduction", Order.Descending}, {"PeriodRunningTotal", Order.Ascending}, {"Period", Order.Ascending}, {"PRCell", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted RES,NWDPRI,MAXPRT,PRT", "ResIndex", 2, 1, Int64.Type)
in
    #"Added Index";

// Query: ResCapIndexLimits
// Purpose: Capture each resource's 2-based start index for the linear running total.
shared ResCapIndexLimits = let
    Source = ResCapPrioritisedTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"StartResIndex", each List.Min([ResIndex]), Int64.Type}})
in
    #"Grouped Rows";

// Query: Resource Running total
// Purpose: Apply the resource cap with one linear pass over the deterministic priority order.
shared #"Resource Running total" = let
    Source = Table.NestedJoin(ResCapPrioritisedTABLE, {"Resource"}, ResCapIndexLimits, {"Resource"}, "ResCapLimits", JoinKind.LeftOuter),
    #"Expanded ResLimits" = Table.ExpandTableColumn(Source, "ResCapLimits", {"StartResIndex"}, {"StartResIndex"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResLimits",{{"ResIndex", Order.Ascending}}),
    #"Added Prefix Count" = Table.AddColumn(#"Sorted Rows", "ResourcePrefixCount", each [ResIndex] - [StartResIndex] + 1, Int64.Type),
    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Added Prefix Count", "Reduction", "ResIndex", "StartResIndex", "ResourcePrefixCount", "ResRunningTotal"),
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "MaxResReduction", each List.Min({[#"ResAv-C'"]})),
    #"Added KEEP" = Table.AddColumn(#"Inserted MAXRESTOTAL", "KeepRunningTotal", each if [ResRunningTotal] > [MaxResReduction] then "Remove for Res" else if [PeriodRunningTotal] > [MaxPeriodReduction] then "Remove for Perod" else if [ResRunningTotal] <= [MaxResReduction] then "Keep" else "Split"),
    #"Filtered KEEP" = Table.SelectRows(#"Added KEEP", each ([KeepRunningTotal] = "Keep")),
    #"Multiplied Column" = Table.TransformColumns(#"Filtered KEEP", {{"Reduction", each _ * -1, type number}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Multiplied Column",{"Resource", "Period", "PRCell", "Reduction", "ResRunningTotal"})
in
    #"Removed Other Columns";

// Query: CheckCap
// Purpose: Retain the A2 resource-cap reconciliation as a connection-only diagnostic.
shared CheckCap = let
    Source = #"Resource Running total",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"MAXResRunningTotal", each List.Max([ResRunningTotal]), type number}, {"AvailabilityReduction", each List.Sum([Reduction]), type number}})
in
    #"Grouped Rows";

// Query: ResPeriodCappedAvailability(C#)TABLE
// Purpose: Apply accepted cap reductions to the authoritative Resource x Period skeleton inside B.
// Notes: Role comes from the skeleton; row order must never determine Role assignment.
shared #"ResPeriodCappedAvailability(C#)TABLE" = let
    Source = #"IMPORT ResPeriodAvailabilityTABLE",
    #"Merged PREMPTY" = Table.NestedJoin(Source, {"Role", "Resource", "Period"}, #"IMPORT ResourcePeriodTABLE_empty", {"Role", "Resource", "Period"}, "ResourcePeriodTABLE_empty", JoinKind.RightOuter),
    #"Removed Source Keys" = Table.RemoveColumns(#"Merged PREMPTY",{"Role", "Resource", "Period"}),
    #"Expanded Authoritative Skeleton" = Table.ExpandTableColumn(#"Removed Source Keys", "ResourcePeriodTABLE_empty", {"Role", "Resource", "Period"}, {"Role", "Resource", "Period"}),
    #"Merged Resource Cap" = Table.NestedJoin(#"Expanded Authoritative Skeleton", {"Resource", "Period"}, #"Resource Running total", {"Resource", "Period"}, "ResourceCap", JoinKind.LeftOuter),
    #"Expanded Resource Cap" = Table.ExpandTableColumn(#"Merged Resource Cap", "ResourceCap", {"Reduction"}, {"Reduction"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Resource Cap",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    #"Merged Roster Status" = Table.NestedJoin(#"Sorted Rows", {"Resource", "Period"}, #"IMPORT ResPeriodWDTABLE", {"Resource", "Period"}, "ResPeriodWDTABLE", JoinKind.FullOuter),
    #"Expanded Roster Status" = Table.ExpandTableColumn(#"Merged Roster Status", "ResPeriodWDTABLE", {"RosteredPeriodStatus"}, {"RosteredPeriodStatus"}),
    #"Replaced Null Reduction" = Table.ReplaceValue(#"Expanded Roster Status",null,0,Replacer.ReplaceValue,{"Reduction"}),
    #"Merged Multi-Period Keep" = Table.NestedJoin(#"Replaced Null Reduction", {"Resource", "Period"}, #"IMPORT MultiPeriod - Remove", {"Resource", "AvailablePeriod"}, "MultiDayPeriod-Remove", JoinKind.LeftOuter),
    #"Expanded Multi-Period Keep" = Table.ExpandTableColumn(#"Merged Multi-Period Keep", "MultiDayPeriod-Remove", {"NoAllocationKeep"}, {"NoAllocationKeep"}),
    // Preserve the accepted A2 business semantics; fractional availability and Split behaviour remain unchanged.
    #"Inserted Capped Availability" = Table.AddColumn(#"Expanded Multi-Period Keep", "AvailabilityCapped", each if (
        [RosteredPeriodStatus] = "RosteredDay"
        or [RosteredPeriodStatus] = "UnavailableAllocatedPeriod"
        or [Availability] = null
        or [NoAllocationKeep] = false
        )
        then 0
        else [Availability] + [Reduction]),
    #"Removed Working Columns" = Table.RemoveColumns(#"Inserted Capped Availability",{"Availability", "Reduction", "NoAllocationKeep"})
in
    #"Removed Working Columns";

// Query: A2_COMPLETE_GRID_CHECK
// Purpose: Reconcile the merged capped output keys to the authoritative Resource x Period skeleton.
// Output: Connection-only diagnostic rows; every failure blocks the published C# interface.
shared A2_COMPLETE_GRID_CHECK = let
    SkeletonKeys = Table.Buffer(Table.Distinct(Table.SelectColumns(#"IMPORT ResourcePeriodTABLE_empty", {"Role", "Resource", "Period"}))),
    ResultKeys = Table.Buffer(Table.SelectColumns(#"ResPeriodCappedAvailability(C#)TABLE", {"Role", "Resource", "Period"})),
    DistinctResultKeys = Table.Distinct(ResultKeys),
    MissingKeys = Table.RowCount(Table.NestedJoin(SkeletonKeys, {"Role", "Resource", "Period"}, DistinctResultKeys, {"Role", "Resource", "Period"}, "Result", JoinKind.LeftAnti)),
    ExtraKeys = Table.RowCount(Table.NestedJoin(DistinctResultKeys, {"Role", "Resource", "Period"}, SkeletonKeys, {"Role", "Resource", "Period"}, "Skeleton", JoinKind.LeftAnti)),
    DuplicateResultKeys = Table.RowCount(ResultKeys) - Table.RowCount(DistinctResultKeys),
    UnexpectedRoleRows = Table.RowCount(Table.SelectRows(ResultKeys, each [Role] <> OutputRole or [Role] = null)),
    Checks = #table(
        type table [Check = text, Status = text, Failures = Int64.Type],
        {
            {"Missing Resource x Period keys", if MissingKeys = 0 then "Pass" else "Fail", MissingKeys},
            {"Extra Resource x Period keys", if ExtraKeys = 0 then "Pass" else "Fail", ExtraKeys},
            {"Duplicate Resource x Period keys", if DuplicateResultKeys = 0 then "Pass" else "Fail", DuplicateResultKeys},
            {"Published Role equals OutputRole", if UnexpectedRoleRows = 0 then "Pass" else "Fail", UnexpectedRoleRows}
        }
    )
in
    Checks;

[ Description = "BUFFER-All cells capped availability" ]
// Query: ResPeriodAvailabilityCapped(C#)TABLE
// Purpose: Publish B's existing C# compatibility interface after merged A2 role and grid validation.
shared #"ResPeriodAvailabilityCapped(C#)TABLE" = let
    RoleFailures = Table.SelectRows(RoleContext_CHECK, each [Status] <> "Pass"),
    GridFailures = Table.SelectRows(A2_COMPLETE_GRID_CHECK, each [Status] <> "Pass"),
    CheckedSource = if Table.IsEmpty(RoleFailures) and Table.IsEmpty(GridFailures) then #"ResPeriodCappedAvailability(C#)TABLE"
        else error Error.Record(
            "CapacityDistribB.MergedCSharpValidation",
            "Role context or Resource x Period completeness failed. Review RoleContext_CHECK and A2_COMPLETE_GRID_CHECK.",
            [RoleFailures = RoleFailures, GridFailures = GridFailures]
        ),
    NarrowSource = Table.RemoveColumns(CheckedSource,{"RosteredPeriodStatus"}),
    BUFFER = Table.Buffer(NarrowSource),
    #"Renamed Columns" = Table.RenameColumns(BUFFER,{{"AvailabilityCapped", "C#"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([#"C#"] <> null)),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows",0,null,Replacer.ReplaceValue,{"C#"}),
    // C# is calculated in this lane, so its published labels must match OutputRole.
    #"Filtered Output Role" = Table.SelectRows(#"Replaced Value", each [Role] = OutputRole)
in
    #"Filtered Output Role";

shared #"ResPeriodAvailabilityCapped(C#)MATRIX" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Removed Columns" = Table.RemoveColumns(Source,{"Role"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "C#", List.Sum)
in
    #"Pivoted Column";

shared ResPeriodAvailabilityTABLE = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE"
in
    Source;

// Query: CapacityDistribB_INPUT_CHECK
// Purpose: Check the required keys of the four inputs used by B's joins before publishing final capacity.
// Output: Four diagnostic rows; a missing column, key error, missing key or duplicate key blocks C###TABLE B.
// Notes: Check the existing consumer inputs, including retained zero-to-null availability rows; do not change their filters.
shared CapacityDistribB_INPUT_CHECK = let
    CheckKeys = (inputName as text, inputTable as table, keys as list) as record =>
        let
            MissingColumns = List.Difference(keys, Table.ColumnNames(inputTable)),
            HasRequiredColumns = List.IsEmpty(MissingColumns),
            // Only key columns are buffered and scanned; the calculation's input table is not altered.
            KeyRows = if HasRequiredColumns then Table.Buffer(Table.SelectColumns(inputTable, keys)) else #table(keys, {}),
            ErrorKeyRows = if HasRequiredColumns then Table.RowCount(Table.SelectRowsWithErrors(KeyRows, keys)) else null,
            KeysWithoutErrors = Table.RemoveRowsWithErrors(KeyRows, keys),
            HasMissingKey = (row as record) as logical => List.AnyTrue(List.Transform(
                Record.FieldValues(row),
                (keyValue) => keyValue = null or (Value.Is(keyValue, type text) and Text.Trim(keyValue) = "")
            )),
            MissingKeyRows = if HasRequiredColumns then Table.RowCount(Table.SelectRows(KeysWithoutErrors, HasMissingKey)) else null,
            CompleteKeys = Table.SelectRows(KeysWithoutErrors, each not HasMissingKey(_)),
            KeyCounts = Table.Group(CompleteKeys, keys, {{"KeyCount", each Table.RowCount(_), Int64.Type}}),
            DuplicateKeyGroups = if HasRequiredColumns then Table.RowCount(Table.SelectRows(KeyCounts, each [KeyCount] > 1)) else null
        in [
            Input = inputName,
            Keys = Text.Combine(keys, ", "),
            MissingColumns = Text.Combine(MissingColumns, ", "),
            ErrorKeyRows = ErrorKeyRows,
            MissingKeyRows = MissingKeyRows,
            DuplicateKeyGroups = DuplicateKeyGroups,
            Passed = HasRequiredColumns and ErrorKeyRows = 0 and MissingKeyRows = 0 and DuplicateKeyGroups = 0
        ],
    Checks = {
        CheckKeys("PeriodDemandTABLE", PeriodDemandTABLE, {"Period"}),
        CheckKeys("ResPeriodAllocationTABLE", ResPeriodAllocationTABLE, {"Resource", "Period"}),
        CheckKeys("ResPeriodAvailabilityCapped(C#)TABLE", #"ResPeriodAvailabilityCapped(C#)TABLE", {"Resource", "Period"}),
        CheckKeys("IMPORT ResPeriodNWDTABLE", #"IMPORT ResPeriodNWDTABLE", {"Resource", "Period"})
    },
    TypedChecks = Table.TransformColumnTypes(Table.FromRecords(Checks), {
        {"Input", type text}, {"Keys", type text}, {"MissingColumns", type text},
        {"ErrorKeyRows", Int64.Type}, {"MissingKeyRows", Int64.Type}, {"DuplicateKeyGroups", Int64.Type}, {"Passed", type logical}
    })
in
    // Reuse these four scalar diagnostic rows within the output's validation and failure reporting.
    Table.Buffer(TypedChecks);

shared #"ResPeriodAvailabilityCapped(C#)SUM" = let
    Source = #"ResPeriodAvailabilityTABLE",
    #"C#" = Source[#"C#"],
    #"Calculated Sum" = List.Sum(#"C#")
in
    #"Calculated Sum";

shared #"PeriodOverallocatedExcess%" = let
    Source = #"ResPeriodMaxC##Reduction",
    #"PERIOD C= Grouped PERIOD.C###" = Table.Group(Source, {"Period"}, {{"PeriodAvailability", each List.Sum([#"C##"]), type nullable number}}),
    #"Merged Queries" = Table.NestedJoin(#"PERIOD C= Grouped PERIOD.C###", {"Period"}, #"PeriodA-DPos (Overallocation)", {"Period"}, "PeriodA-DPos (Excess)", JoinKind.LeftOuter),
    #"Expanded PeriodA-DPos (Excess)" = Table.ExpandTableColumn(#"Merged Queries", "PeriodA-DPos (Excess)", {"D", "A-D"}, {"D", "A-D"}),
    #"Added EXCESS%" = Table.AddColumn(#"Expanded PeriodA-DPos (Excess)", "Excess%", each [#"A-D"] / [D]),
    #"Filtered NULL" = Table.SelectRows(#"Added EXCESS%", each [#"Excess%"] <> null),
    #"Sorted Rows1" = Table.Sort(#"Filtered NULL",{{"Excess%", Order.Descending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows1";

shared #"PeriodA-DPos (Overallocation)" = let
    Source = #"PeriodA-D",
    #"Inserted Subtraction" = Table.AddColumn(Source, "A-D", each [A]-[D], type number),
    #"Filtered Rows" = Table.SelectRows(#"Inserted Subtraction", each [#"A-D"] > 0)
in
    #"Filtered Rows";

[ Description = "BUFFER" ]
// Query: PrioritiseReductionAvail-Setup
// Purpose: Attach resource-period priorities and retain the existing reduction sort and index.
shared #"PrioritiseReductionAvail-Setup" = let
    Source = Table.NestedJoin(#"ResPeriodMaxC##Reduction", {"Period"}, #"PeriodOverallocatedExcess%", {"Period"}, "PeriodOverallocatedPriroristised", JoinKind.LeftOuter),
    #"Expanded PeriodOverallocatedPriroristised" = Table.ExpandTableColumn(Source, "PeriodOverallocatedPriroristised", {"Excess%"}, {"Excess%"}),
    // Pair Resource with Resource and Period with Period; positional join keys must use the same order.
    #"Merged Queries" = Table.NestedJoin(#"Expanded PeriodOverallocatedPriroristised", {"Resource", "Period"}, #"IMPORT ResPeriodNWDTABLE", {"Resource", "Period"}, "IMPORT ResPeriodNWDTABLE", JoinKind.LeftOuter),
    #"Expanded IMPORT ResPeriodNWDTABLE" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ResPeriodNWDTABLE", {"NWDPriority"}, {"NWDPriority"}),
    #"Inserted MINAVAILABLEREDUCTION" = Table.AddColumn(#"Expanded IMPORT ResPeriodNWDTABLE", "MinimumAvailable", each List.Min({[#"C##"], [#"A-D"]})),
    // Resource is the stable tie-breaker inside an otherwise-equal Period priority.
    #"Sorted Rows" = Table.Sort(#"Inserted MINAVAILABLEREDUCTION",{{"Period", Order.Ascending}, {"Excess%", Order.Descending}, {"ResPeriodC##MAXReduction", Order.Descending}, {"Resource", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRows", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

shared #"PeriodIndexLimits-ve (C##-D,C##)" = let
    Source = #"PrioritiseReductionAvail-Setup",
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([IndexAllRows]), type number}, {"PeriodExcessAvailable", each List.Max([#"ResPeriodC##MAXReduction"]), type number}, {"PeriodAvailableNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

// Query: SubtractOverallatedPeriods
// Purpose: Apply period reduction limits in the existing indexed priority order.
shared SubtractOverallatedPeriods = let
    Source = Table.NestedJoin(#"PrioritiseReductionAvail-Setup", {"Period"}, #"PeriodIndexLimits-ve (C##-D,C##)", {"Period"}, "ResLimits-Subtract", JoinKind.LeftOuter),
    #"Expanded ResLimits-Subtract" = Table.ExpandTableColumn(Source, "ResLimits-Subtract", {"StartPeriodIndex", "PeriodExcessAvailable", "PeriodAvailableNumber"}, {"StartPeriodIndex", "PeriodExcessAvailable", "PeriodAvailableNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResLimits-Subtract",{{"IndexAllRows", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction", each [IndexAllRows] - [StartPeriodIndex]+1),
    #"Reordered Columns" = Table.ReorderColumns(#"Inserted Subtraction",{"Period", "Resource", "C##", "ExcessC##-D", "C##/D", "A-D", "ResPeriodC##MAXReduction", "Excess%", "NWDPriority", "IndexAllRows", "StartPeriodIndex", "PeriodExcessAvailable", "MinimumAvailable", "PeriodAvailableNumber", "Subtraction"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"Subtraction", Int64.Type}}),




    #"Added RUNNINGRESTOTAL" = fnAddIndexedRunningTotal(#"Changed Type", "MinimumAvailable", "IndexAllRows", "StartPeriodIndex", "Subtraction", "RunningPeriodTotal"),
    #"Added KEEP" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "KeepRunningTotal", each if [RunningPeriodTotal] > [#"ResPeriodC##MAXReduction"] then "Remove" else if [RunningPeriodTotal] <= [#"ResPeriodC##MAXReduction"] then "Keep" else "Split")
in
    #"Added KEEP";

shared #"ResIndex-Steup" = let
    Source = SubtractOverallatedPeriods,
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource"}, #"ResRosterAvailabilityC##CapReduction", {"Resource"}, "ResRosterAvailabilityC##CapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityC##CapReduction" = Table.ExpandTableColumn(#"Merged Queries", "ResRosterAvailabilityC##CapReduction", {"MaxC##Reduction"}, {"MaxC##Reduction"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Expanded ResRosterAvailabilityC##CapReduction",{"Period", "Resource", "A-D", "ResPeriodC##MAXReduction", "MinimumAvailable", "RunningPeriodTotal", "KeepRunningTotal", "MaxC##Reduction"}),
    // Period is the stable tie-breaker inside an otherwise-equal Resource priority.
    #"Sorted Rows" = Table.Sort(#"Removed Other Columns",{{"Resource", Order.Ascending}, {"MaxC##Reduction", Order.Descending}, {"Period", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRows", 2, 1, Int64.Type)
in
    #"Added Index";

shared ResIndexLimits = let
    Source = #"ResIndex-Steup",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"StartResIndex", each List.Min([IndexAllRows]), type number}, {"ResExcessAvailable", each List.Max([#"ResPeriodC##MAXReduction"]), type number}, {"ResAvailableNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

// Query: SubtractOverallocatedResources
// Purpose: Retain reduction rows within both resource and period caps using the existing indexed order.
shared SubtractOverallocatedResources = let
    Source = Table.NestedJoin(#"ResIndex-Steup", {"Resource"}, ResIndexLimits, {"Resource"}, "ResIndexLimits", JoinKind.LeftOuter),
    #"Expanded ResIndexLimits" = Table.ExpandTableColumn(Source, "ResIndexLimits", {"StartResIndex", "ResExcessAvailable", "ResAvailableNumber"}, {"StartResIndex", "ResExcessAvailable", "ResAvailableNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResIndexLimits",{{"IndexAllRows", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction", each [IndexAllRows] - [StartResIndex]+1),
    #"Added Custom" = fnAddIndexedRunningTotal(#"Inserted Subtraction", "MinimumAvailable", "IndexAllRows", "StartResIndex", "Subtraction", "RunningResTotal"),
    #"Replaced Value" = Table.ReplaceValue(#"Added Custom",null,0,Replacer.ReplaceValue,{"A-D"}),
    #"Added KEEP RES" = Table.AddColumn(#"Replaced Value", "KeepRes", each if [RunningResTotal] <= [#"MaxC##Reduction"] then "Keep" else "Remove"),
    #"Added Conditional Column" = Table.AddColumn(#"Added KEEP RES", "KeepPR", each if [KeepRunningTotal] <> "Keep" then null else if [KeepRes] <> "Keep" then null else "Keep"),
    #"Filtered Rows" = Table.SelectRows(#"Added Conditional Column", each ([KeepPR] = "Keep"))
in
    #"Filtered Rows";

shared ReduceCheck = let
    Source = SubtractOverallocatedResources,
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"MaxReDistrib", each List.Max([RunningResTotal]), type number}, {"Available", each List.Sum([MinimumAvailable]), type number}})
in
    #"Grouped Rows";

shared #"C##TABLE" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, #"ReDistribPeriodAvailbility TABLE", {"Resource", "Period"}, "ReDistribAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ReDistribAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ReDistribAvailabilityTABLE", {"C#"}, {"C#.1"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded ReDistribAvailabilityTABLE",null,0,Replacer.ReplaceValue,{"C#.1"}),
    #"CHANGE C" = Table.AddColumn(#"Replaced Value", "C##", each [#"C#"] - [#"C#.1"], type number),
    #"Removed Columns" = Table.RemoveColumns(#"CHANGE C",{"C#", "C#.1"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Removed Columns",0,null,Replacer.ReplaceValue,{"C##"})
in
    #"Replaced Value1";

shared #"C##SUM" = let
    Source = #"C##TABLE",
    #"C##" = Source[#"C##"],
    #"Calculated Sum" = List.Sum(#"C##")
in
    #"Calculated Sum";

shared #"C##MATRIX" = let
    Source = #"C##TABLE",
    #"Sorted Rows" = Table.Sort(Source,{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column1" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "C##", List.Sum)
in
    #"Pivoted Column1";

shared #"PeriodC#-D" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Grouped PERIODC#" = Table.Group(Source, {"Period"}, {{"PeriodC#", each List.Sum([#"C#"]), type nullable number}}),
    #"Merged Queries" = Table.NestedJoin(#"Grouped PERIODC#", {"Period"}, PeriodDemandTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C#-D", each [#"PeriodC#"] - [D], type number),
    #"Inserted Division" = Table.AddColumn(#"Inserted Subtraction", "Division", each [#"PeriodC#"] / [D], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Division",{{"Division", "C#/D"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"D"})
in
    #"Removed Columns";

shared #"PeriodC#-DPos (Excess)TABLE" = let
    Source = #"PeriodC#-D",
    #"Filtered Rows" = Table.SelectRows(Source, each [#"C#-D"] > 0),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"C#-D", "ExcessC#-D"}})
in
    #"Renamed Columns";

shared PeriodCapacityTABLE = let
    Source = ResPeriodAvailabilityTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"PeriodCapacity", each List.Sum([#"C#"]), type nullable number}})
in
    #"Grouped Rows";

shared ResMaxAvailability = let
    Source = #"IMPORT AvailabilityOriginal",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"ResAvailability", each List.Sum([Availability]), type nullable number}}),
    // ResMaxAvail is a Resource measure, so attach the contract cap after Resource aggregation.
    #"Merged Resource Contract" = Table.NestedJoin(#"Grouped Rows", {"Resource"}, BResourceContract, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    #"Expanded Effective Shift Cap" = Table.ExpandTableColumn(#"Merged Resource Contract", "ResourceContract", {"Effective Shift Cap"}, {"Effective Shift Cap"}),
    #"Added RESMAXAVAILABILITY" = Table.AddColumn(#"Expanded Effective Shift Cap", "ResMaxAvail", each
        // A Resource with no original availability does not need a contract cap in this calculation.
        if [ResAvailability] = null then 0
        else if [ResAvailability] = 0 then 0
        else if [Effective Shift Cap] = null then error Error.Record(
            "CapacityDistribB.MissingResourceCap",
            "A Resource with positive original availability has no Effective Shift Cap.",
            [Resource = [Resource], ResAvailability = [ResAvailability]]
        )
        else if [ResAvailability] > [Effective Shift Cap] then [Effective Shift Cap]
        else [ResAvailability], type number),
    #"Removed Effective Shift Cap" = Table.RemoveColumns(#"Added RESMAXAVAILABILITY", {"Effective Shift Cap"})
in
    #"Removed Effective Shift Cap";

// Query: ResourcesLIST
// Purpose: Publish the narrow distinct Resource dimension used by the Resource x Period grid.
shared ResourcesLIST = let
    Source = ResPeriodAvailabilityTABLE,
    #"Removed Duplicates" = Table.Distinct(Table.SelectColumns(Source, {"Resource"}))
in
    #"Removed Duplicates";

// Query: PeriodLIST
// Purpose: Publish the narrow distinct Period dimension before the Cartesian Resource x Period expansion.
shared PeriodLIST = let
    Source = PeriodDemandTABLE,
    #"Removed Duplicates" = Table.Distinct(Table.SelectColumns(Source, {"Period"}))
in
    #"Removed Duplicates";

shared #"PeriodC##CapacityTABLE" = let
    Source = #"C##TABLE",
    #"Grouped PERIODC##" = Table.Group(Source, {"Period"}, {{"PeriodC##", each List.Sum([#"C##"]), type nullable number}})
in
    #"Grouped PERIODC##";

shared #"PeriodC##-D" = let
    Source = #"PeriodC##CapacityTABLE",
    #"Merged Queries" = Table.NestedJoin(Source, {"Period"}, PeriodDemandTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C##-D", each [#"PeriodC##"] - [D], type number),
    #"Inserted Division" = Table.AddColumn(#"Inserted Subtraction", "Division", each [#"PeriodC##"] / [D], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted Division",{{"Division", "C##/D"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"D"})
in
    #"Removed Columns";

shared #"PeriodC##-DPos (Excess)TABLE" = let
    Source = #"PeriodC##-D",
    #"Filtered Rows" = Table.SelectRows(Source, each [#"C##-D"] > 0),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"C##-D", "ExcessC##-D"}})
in
    #"Renamed Columns";

shared #"PeriodMaxReduction - Empty Tables" = let
    Source = Table.NestedJoin(PeriodLIST, {"Period"}, #"PeriodC##-DPos (Excess)TABLE", {"Period"}, "PeriodC##-DPos (Excess)TABLE", JoinKind.LeftOuter),
    #"Expanded PeriodC##-DPos (Excess)TABLE" = Table.ExpandTableColumn(Source, "PeriodC##-DPos (Excess)TABLE", {"ExcessC##-D", "C##/D"}, {"ExcessC##-D", "C##/D"})
in
    #"Expanded PeriodC##-DPos (Excess)TABLE";

shared #"PeriodC##MaxReduction" = let
    Source = #"PeriodMaxReduction - Empty Tables",
    #"Filtered Rows" = Table.SelectRows(Source, each ([#"ExcessC##-D"] <> null)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Period"}, #"PeriodA-DPos (Overallocation)", {"Period"}, "PeriodA-D", JoinKind.LeftOuter),
    #"Expanded PeriodA-D" = Table.ExpandTableColumn(#"Merged Queries", "PeriodA-D", {"A-D"}, {"A-D"}),
    #"Inserted Minimum" = Table.AddColumn(#"Expanded PeriodA-D", "PeriodMaxC##Reduction", each List.Min({ [#"A-D"]}))
in
    #"Inserted Minimum";

[ Description = "BUFFER" ]
shared #"ResPeriodMaxC##Reduction" = let
    Source = Table.NestedJoin(#"PeriodC##MaxReduction", {"Period"}, #"C##TABLE", {"Period"}, "C##TABLE", JoinKind.LeftOuter),
    #"Expanded C##TABLE" = Table.ExpandTableColumn(Source, "C##TABLE", {"Resource", "C##"}, {"Resource", "C##"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded C##TABLE", each ([#"C##"] <> null)),
    #"Merged Queries1" = Table.NestedJoin(#"Filtered Rows", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Filtered NOT ALLOCATED" = Table.SelectRows(#"Expanded ResPeriodAllocationTABLE", each ([Allocation] = null)),
    #"Reordered Columns" = Table.ReorderColumns(#"Filtered NOT ALLOCATED",{"Period", "Resource", "C##", "PeriodMaxC##Reduction","ExcessC##-D", "C##/D", "A-D"}),
    #"Merged Queries" = Table.NestedJoin(#"Reordered Columns", {"Resource"}, #"ResRosterAvailabilityC##CapReduction", {"Resource"}, "ResPeriodOverallocationC##Reduction", JoinKind.LeftOuter),
    #"Expanded ResPeriodOverallocationC##Reduction" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodOverallocationC##Reduction", {"MaxC##Reduction"}, {"MaxC##Reduction"}),
    #"Filtered MAXREDUCTION C## <>0 or NULL" = Table.SelectRows(#"Expanded ResPeriodOverallocationC##Reduction", each ([#"MaxC##Reduction"] <> null) and ([#"C##"] <> 0)),
    // Period or res max reduction
    #"Inserted Minimum" = Table.AddColumn(#"Filtered MAXREDUCTION C## <>0 or NULL", "ResPeriodC##MAXReduction", each List.Min({[#"PeriodMaxC##Reduction"], [#"MaxC##Reduction"]}), type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Minimum",{"PeriodMaxC##Reduction", "MaxC##Reduction"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER;

// Query: C###TABLE B
// Purpose: Publish final B capacity only after the required input join keys pass validation.
// Output: Preserve the existing final-capacity table interface used by C###SUM and C###MATRIX.
shared #"C###TABLE B" = let
    InputChecks = CapacityDistribB_INPUT_CHECK,
    ContractChecks = BResourceContract_CHECK,
    RoleChecks = RoleContext_CHECK,
    // Gate the source actually consumed below so lazy evaluation cannot skip required validation.
    Source = if List.AllTrue(InputChecks[Passed])
        and List.AllTrue(List.Transform(ContractChecks[Status], each _ = "Pass"))
        and List.AllTrue(List.Transform(RoleChecks[Status], each _ = "Pass"))
        then #"C##TABLE"
        else error Error.Record(
            "CapacityDistribB.InputValidation",
            "Required join keys, Resource contracts or role context failed validation. Review the B check queries.",
            [
                InputFailures = Table.SelectRows(InputChecks, each [Passed] <> true),
                ContractFailures = Table.SelectRows(ContractChecks, each [Status] <> "Pass"),
                RoleFailures = Table.SelectRows(RoleChecks, each [Status] <> "Pass")
            ]
        ),
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, SubtractOverallocatedResources, {"Resource", "Period"}, "SubtractAvailablilityTABLE", JoinKind.LeftOuter),
    #"Expanded SubtractAvailablilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "SubtractAvailablilityTABLE", {"MinimumAvailable", "KeepPR"}, {"MinimumAvailable", "KeepPR"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded SubtractAvailablilityTABLE",null,0,Replacer.ReplaceValue,{"MinimumAvailable"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Replaced Value", "C###", each [#"C##"] - [MinimumAvailable], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"C##", "MinimumAvailable", "KeepPR"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Removed Columns",0,null,Replacer.ReplaceValue,{"C###"}),
    #"Merged Queries1" = Table.NestedJoin(#"Replaced Value1", {"Resource"}, ResMaxAvailability, {"Resource"}, "ResMaxAvailability", JoinKind.LeftOuter),
    #"Expanded ResMaxAvailability" = Table.ExpandTableColumn(#"Merged Queries1", "ResMaxAvailability", {"ResAvailability", "ResMaxAvail"}, {"ResAvailability", "ResMaxAvail"})
in
    #"Expanded ResMaxAvailability";

shared #"C###SUM" = let
    Source = #"C###TABLE B",
    #"C###" = Source[#"C###"],
    #"Calculated Sum" = List.Sum(#"C###")
in
    #"Calculated Sum";

shared #"C###MATRIX" = let
    Source = #"C###TABLE B",
    #"Sorted Rows" = Table.Sort(Source,{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "C###", List.Sum)
in
    #"Pivoted Column";

[ Description = "BUFFER" ]
// Query: ResourcePeriods-EmptyTABLE
// Purpose: Build the complete narrow Resource x Period key grid without carrying Demand attributes into every row.
shared #"ResourcePeriods-EmptyTABLE" = let
    Source = PeriodLIST,
    #"Added Custom" = Table.AddColumn(Source, "Custom", each ResourcesLIST),
    #"Expanded Custom" = Table.ExpandTableColumn(#"Added Custom", "Custom", {"Resource"}, {"Resource"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Custom",{{"Period", Int64.Type}}),
    BUFFER = Table.Buffer(#"Changed Type")
in
    BUFFER;

shared ResPeriodAllocationMATRIX = let
    Source = #"ResourcePeriods-EmptyTABLE",
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResPeriodAllocationTABLE",{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "Allocation", List.Sum)
in
    #"Pivoted Column";

// Query: RolePathTABLE
// Purpose: Resolve this workbook's folder and dynamic role through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl (one-row FilePath table or single named cell) and the public CentriSyncPaths table.
// Output: Existing Variable Name / Value rows used by RolePath and Role.
shared RolePathTABLE =
let
    NormalizePath = (value as nullable text) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(value), "/", "\"), "\"),
    FilePathUrl =
    let
        // Accept the legacy FilePAthUrl casing without masking a missing or malformed input.
        PathInputs = Table.SelectRows(Excel.CurrentWorkbook(), each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
        Source = if Table.RowCount(PathInputs) = 1 then PathInputs{0}[Content]
            else error "Expected exactly one FilePathUrl table or named cell in this workbook.",
        // A single named cell is exposed by Excel as Column1.
        PathColumn = if Table.HasColumns(Source, {"FilePath"}) then Table.SelectColumns(Source, {"FilePath"})
            else if Table.ColumnNames(Source) = {"Column1"} then Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must contain a FilePath column or be a single named cell.",
        ValidatedRows = if Table.RowCount(PathColumn) = 1 then PathColumn
            else error "FilePathUrl must contain exactly one data row.",
        TypedPath = Table.TransformColumnTypes(ValidatedRows, {{"FilePath", type text}}),
        BufferedTable = Table.Buffer(TypedPath)
    in
        BufferedTable,

    RawFilePath = NormalizePath(FilePathUrl{0}[FilePath]),
    NonBlankFilePath = if RawFilePath = null or RawFilePath = "" then
        error "FilePathUrl is blank. Save this workbook and recalculate its CELL filename formula."
        else RawFilePath,
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet; retain only the workbook path.
    WorkbookPath = if Text.Contains(NonBlankFilePath, "[") then
        Text.BeforeDelimiter(NonBlankFilePath, "[") & Text.BetweenDelimiters(NonBlankFilePath, "[", "]")
        else NonBlankFilePath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "CapacityDistrib(B)-shifts.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Use =CELL(""filename"",A1) in its input cell, then save and recalculate CapacityDistrib(B)-shifts.xlsx.",
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
    // Match full roots for both SharePoint and local paths; no Site-name join is required.
    SharePointCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SharepointRootUrl], type text),
    SharePointDocumentsCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\Shared Documents", type text),
    LocalCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SyncedFolderRootPath], type text),
    MatchCandidates = Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
    MatchCandidates_WithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(
        MatchCandidates_WithLength,
        each [MatchRoot] <> null
            and Text.Trim([MatchRoot]) <> ""
            and [SyncedFolderRootPath] <> null
            and Text.Trim([SyncedFolderRootPath]) <> ""
            and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
            // A root must end at a folder boundary, not part-way through a different folder name.
            and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    LongestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0}
        else error "FilePathUrl has no usable CentriSyncPaths mapping. Check SharepointRootUrl and SyncedFolderRootPath for: " & FilePath,
    // Duplicate matching rows may agree, but conflicting destinations must not depend on row order.
    BestMatches = Table.SelectRows(SortedMatches, each [MatchRootLength] = LongestMatch[MatchRootLength]),
    BestMatch = if Table.IsEmpty(SortedMatches) then LongestMatch
        else if List.Count(List.Distinct(BestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase)) = 1 then LongestMatch
        else error "CentriSyncPaths contains conflicting local folders for: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    FolderSegments = List.Select(Text.Split(WorkbookFolder, "\"), each _ <> ""),
    UnitsIndex = List.PositionOf(FolderSegments, "UNITS", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    CalcIndex = List.PositionOf(FolderSegments, "2. Calculations", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    // Imports require UNITS/<unit>/2. Calculations/<role>; unit and role names remain dynamic.
    RootPath = if UnitsIndex >= 0
        and CalcIndex = UnitsIndex + 2
        and List.Count(FolderSegments) = CalcIndex + 2
        and Text.Trim(FolderSegments{UnitsIndex + 1}) <> ""
        and Text.Trim(FolderSegments{CalcIndex + 1}) <> "" then WorkbookFolder
        else error "FilePathUrl must identify this workbook under UNITS\<unit>\2. Calculations\<role>. Save and recalculate its path input.",
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    // Preserve the role folder name so copies into another role folder use that role automatically.
    Role = Segments{CalcIndex + 1},
    FileName = InputFileName,
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", RootPath},
            {"Client", Client},
            {"Date", Date},
            {"Unit", Unit},
            {"Role", Role},
            {"FileName", FileName}
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

shared RolePath = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    Folder = #"Renamed Columns"{0}[Folder]
in
    Folder;
