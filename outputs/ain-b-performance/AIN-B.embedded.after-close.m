section Section1;


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
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"Resource", Order.Ascending}, {"C#/D", Order.Descending}, {"ExcessC#-D", Order.Descending}, {"NWDPriority", Order.Descending}}),
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

shared ReDistributeResAvailability = let
    Source = Table.NestedJoin(#"PrioritiseRedistribAvail-Distrib", {"Resource"}, #"ResIndexLimits (ResMaxReduction)", {"Resource"}, "REsLimits", JoinKind.LeftOuter),
    #"Expanded REsLimits" = Table.ExpandTableColumn(Source, "REsLimits", {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}, {"StartResIndex", "ResExcessAvailable", "ResAvailNumber"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded REsLimits", {"Resource", "Period"}, #"ResPeriodAvailabilityCapped(C#)TABLE", {"Resource", "Period"}, "ResPeriodAvailabilityCapped(C#)TABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAvailabilityCapped(C#)TABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAvailabilityCapped(C#)TABLE", {"C#"}, {"C#"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded ResPeriodAvailabilityCapped(C#)TABLE",{{"IndexAllRowPrioritySort", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows1", "Subtraction", each [IndexAllRowPrioritySort] - [StartResIndex]+1),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Subtraction",{{"Subtraction", Int64.Type}}),
    #"Added RUNNINGRESTOTAL" = Table.AddColumn(#"Changed Type", "RunningResTotal", each List.Sum(List.Range(#"Changed Type" [#"C#"], [StartResIndex]-2
, [Subtraction]))),
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
shared #"ReDistribPeriodAvailbility TABLE" = let
    Source = Table.NestedJoin(ReDistribResAvailabilityTABLE, {"Period"}, #"PeriodIndexLimits (ExcessC#-D)", {"Period"}, "PeriodIndexLimits", JoinKind.LeftOuter),
    #"Expanded PeriodIndexLimits" = Table.ExpandTableColumn(Source, "PeriodIndexLimits", {"StartPeriodIndex"}, {"StartPeriodIndex"}),
    #"Added Custom" = Table.AddColumn(#"Expanded PeriodIndexLimits", "ApplicableC#", each if [KeepRunningTotal] = "Remove" then 0 else [#"C#"]),
    #"Sorted INDEX" = Table.Sort(#"Added Custom",{{"Index", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted INDEX", "Subtraction", each [Index] - [StartPeriodIndex]+1, type number),
    #"Added PERIODRUNNINGTOTAL" = Table.AddColumn(#"Inserted Subtraction", "PeriodRunningTotal", each List.Sum(List.Range(#"Sorted INDEX" [#"ApplicableC#"], [StartPeriodIndex]-2
, [Subtraction]))),
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
    #"Added AVAILCAP" = Table.AddColumn(#"Replaced Value", "RosterAvailabilityCAPPED", each if [RosterAvailability] > MaxAvailability
then MaxAvailability
else [RosterAvailability]),
    #"Inserted Subtraction" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"RosterAvailability", "RosterAvailabilityCAPPED"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailabilityReduction", "MaxC##Reduction"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([#"MaxC##Reduction"] <> 0))
in
    #"Filtered Rows";

shared Role = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Role")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

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

shared #"IMPORT Demand" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    Demand_Prepare_Table = Source{[Item="Demand_Prepare",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Demand_Prepare_Table,{{"Shift", type text}, {"Role", type text}, {"Facility", type text}, {"D", Int64.Type}, {"Date", type date}, {"Period", Int64.Type}})
in
    #"Changed Type";

shared #"IMPORT Allocation" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAllocationTABLE_Table = Source{[Item="ResPeriodAllocationTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAllocationTABLE_Table,{{"Role", type text}, {"Allocation", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Period] <> null))
in
    #"Filtered Rows";

[ Description = "BUFFER" ]
shared #"IMPORT AvailabilityOriginal" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    BUFFER = Table.Buffer(#"Changed Type")
in
    BUFFER;

shared #"IMPORT ResPeriodAvailabilityCapped(C#)" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodCappedAvailability_C__TABLE_Table = Source{[Item="ResPeriodCappedAvailability_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodCappedAvailability_C__TABLE_Table,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"AvailabilityCapped", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

shared #"IMPORT ResPeriodNWDTABLE" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodWDTABLE_Sheet = Source{[Item="ResPeriodWDTABLE",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(ResPeriodWDTABLE_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Resource", Int64.Type}, {"Day", Int64.Type}, {"PotentialAvailability", type text}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"RosteredPeriodStatus", type text}})
in
    #"Changed Type";

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
    #"Added AVAILCAP" = Table.AddColumn(#"Replaced Value", "RosterAvailabilityCAPPED", each if [RosterAvailability] > MaxAvailability
then MaxAvailability
else [RosterAvailability]),
    #"Inserted Subtraction" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"RosterAvailability", "RosterAvailabilityCAPPED"})
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
    #"Removed Columns" = Table.RemoveColumns(Source,{"Shift", "Role", "Facility", "Date"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Resource", "Period"}, #"IMPORT AvailabilityOriginal", {"Resource", "Period"}, "IMPORT AvailabilityOriginal", JoinKind.LeftOuter),
    #"Expanded IMPORT AvailabilityOriginal" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT AvailabilityOriginal", {"Availability"}, {"Availability"}),
    #"Sorted Rows" = Table.Sort(#"Expanded IMPORT AvailabilityOriginal",{{"Period", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Sorted Rows", {{"Period", type text}}, "en-AU")[Period]), "Period", "Availability", List.Sum)
in
    #"Pivoted Column";

shared PeriodDemandTABLE = let
    Source = #"IMPORT Demand",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Period", Int64.Type}})
in
    #"Changed Type";

shared PeriodDemandMATRIX = let
    Source = PeriodDemandTABLE,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Period", "D"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Other Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Other Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "D", List.Sum)
in
    #"Pivoted Column";

[ Description = "BUFFER-All cells capped availability" ]
shared #"ResPeriodAvailabilityCapped(C#)TABLE" = let
    Source = #"IMPORT ResPeriodAvailabilityCapped(C#)",
    BUFFER = Table.Buffer(Source),
    #"Renamed Columns" = Table.RenameColumns(BUFFER,{{"AvailabilityCapped", "C#"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([#"C#"] <> null)),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows",0,null,Replacer.ReplaceValue,{"C#"}),
    #"Filtered Rows1" = Table.SelectRows(#"Replaced Value", each ([Role] = Role) )
in
    #"Filtered Rows1";

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
shared #"PrioritiseReductionAvail-Setup" = let
    Source = Table.NestedJoin(#"ResPeriodMaxC##Reduction", {"Period"}, #"PeriodOverallocatedExcess%", {"Period"}, "PeriodOverallocatedPriroristised", JoinKind.LeftOuter),
    #"Expanded PeriodOverallocatedPriroristised" = Table.ExpandTableColumn(Source, "PeriodOverallocatedPriroristised", {"Excess%"}, {"Excess%"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded PeriodOverallocatedPriroristised", {"Period", "Resource"}, #"IMPORT ResPeriodNWDTABLE", {"Resource", "Period"}, "IMPORT ResPeriodNWDTABLE", JoinKind.LeftOuter),
    #"Expanded IMPORT ResPeriodNWDTABLE" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ResPeriodNWDTABLE", {"NWDPriority"}, {"NWDPriority"}),
    #"Inserted MINAVAILABLEREDUCTION" = Table.AddColumn(#"Expanded IMPORT ResPeriodNWDTABLE", "MinimumAvailable", each List.Min({[#"C##"], [#"A-D"]})),
    #"Sorted Rows" = Table.Sort(#"Inserted MINAVAILABLEREDUCTION",{{"Period", Order.Ascending}, {"Excess%", Order.Descending}, {"ResPeriodC##MAXReduction", Order.Descending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRows", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

shared #"PeriodIndexLimits-ve (C##-D,C##)" = let
    Source = #"PrioritiseReductionAvail-Setup",
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([IndexAllRows]), type number}, {"PeriodExcessAvailable", each List.Max([#"ResPeriodC##MAXReduction"]), type number}, {"PeriodAvailableNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
shared SubtractOverallatedPeriods = let
    Source = Table.NestedJoin(#"PrioritiseReductionAvail-Setup", {"Period"}, #"PeriodIndexLimits-ve (C##-D,C##)", {"Period"}, "ResLimits-Subtract", JoinKind.LeftOuter),
    #"Expanded ResLimits-Subtract" = Table.ExpandTableColumn(Source, "ResLimits-Subtract", {"StartPeriodIndex", "PeriodExcessAvailable", "PeriodAvailableNumber"}, {"StartPeriodIndex", "PeriodExcessAvailable", "PeriodAvailableNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResLimits-Subtract",{{"IndexAllRows", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction", each [IndexAllRows] - [StartPeriodIndex]+1),
    #"Reordered Columns" = Table.ReorderColumns(#"Inserted Subtraction",{"Period", "Resource", "C##", "ExcessC##-D", "C##/D", "A-D", "ResPeriodC##MAXReduction", "Excess%", "NWDPriority", "IndexAllRows", "StartPeriodIndex", "PeriodExcessAvailable", "MinimumAvailable", "PeriodAvailableNumber", "Subtraction"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Reordered Columns",{{"Subtraction", Int64.Type}}),




    #"Added RUNNINGRESTOTAL" = Table.AddColumn(#"Changed Type", "RunningPeriodTotal", each List.Sum(List.Range(#"Changed Type" [MinimumAvailable], [StartPeriodIndex]-2
, [Subtraction]))),
    #"Added KEEP" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "KeepRunningTotal", each if [RunningPeriodTotal] > [#"ResPeriodC##MAXReduction"] then "Remove" else if [RunningPeriodTotal] <= [#"ResPeriodC##MAXReduction"] then "Keep" else "Split")
in
    #"Added KEEP";

shared #"ResIndex-Steup" = let
    Source = SubtractOverallatedPeriods,
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource"}, #"ResRosterAvailabilityC##CapReduction", {"Resource"}, "ResRosterAvailabilityC##CapReduction", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailabilityC##CapReduction" = Table.ExpandTableColumn(#"Merged Queries", "ResRosterAvailabilityC##CapReduction", {"MaxC##Reduction"}, {"MaxC##Reduction"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Expanded ResRosterAvailabilityC##CapReduction",{"Period", "Resource", "A-D", "ResPeriodC##MAXReduction", "MinimumAvailable", "RunningPeriodTotal", "KeepRunningTotal", "MaxC##Reduction"}),
    #"Sorted Rows" = Table.Sort(#"Removed Other Columns",{{"Resource", Order.Ascending}, {"MaxC##Reduction", Order.Descending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "IndexAllRows", 2, 1, Int64.Type)
in
    #"Added Index";

shared ResIndexLimits = let
    Source = #"ResIndex-Steup",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"StartResIndex", each List.Min([IndexAllRows]), type number}, {"ResExcessAvailable", each List.Max([#"ResPeriodC##MAXReduction"]), type number}, {"ResAvailableNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

shared SubtractOverallocatedResources = let
    Source = Table.NestedJoin(#"ResIndex-Steup", {"Resource"}, ResIndexLimits, {"Resource"}, "ResIndexLimits", JoinKind.LeftOuter),
    #"Expanded ResIndexLimits" = Table.ExpandTableColumn(Source, "ResIndexLimits", {"StartResIndex", "ResExcessAvailable", "ResAvailableNumber"}, {"StartResIndex", "ResExcessAvailable", "ResAvailableNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ResIndexLimits",{{"IndexAllRows", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction", each [IndexAllRows] - [StartResIndex]+1),
    #"Added Custom" = Table.AddColumn(#"Inserted Subtraction", "RunningResTotal", each List.Sum(List.Range(#"Inserted Subtraction" [#"MinimumAvailable"], [StartResIndex]-2
, [Subtraction]))),
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
    #"Added RESMAXAVAILABILITY" = Table.AddColumn(#"Grouped Rows", "ResMaxAvail", each if [ResAvailability] > MaxAvailability 
then MaxAvailability
else [ResAvailability])
in
    #"Added RESMAXAVAILABILITY";

shared ResourcesLIST = let
    Source = ResPeriodAvailabilityTABLE,
    Resource = Source[Resource],
    #"Removed Duplicates" = List.Distinct(Resource),
    #"Converted to Table" = Table.FromList(#"Removed Duplicates", Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Resource"}})
in
    #"Renamed Columns";

shared PeriodLIST = let
    Source = PeriodDemandTABLE,
    #"Removed Columns" = Table.RemoveColumns(Source,{"D"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns")
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

[ Description = "BUFFER" ]
shared #"C###TABLE B" = let
    Source = #"C##TABLE",
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
    #"Removed Columns" = Table.RemoveColumns(Source,{"Role", "Shift", "Facility", "Date"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
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