// Power Query from: CapacityDistrib(A.1)-shifts.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\AINC4\CapacityDistrib(A.1)-shifts.xlsx
// Extracted: 2026-05-21T00:53:35.669Z

section Section1;

shared #"IMPORT ShiftUnitDemandHRS !!" = let
    Source = Excel.Workbook(File.Contents(FilePath&"\2. Calculations\Demand.xlsx"), null, true),
    ShiftDemandHCAverageANACC_Table = Source{[Item="ShiftDemandHCAverageANACC",Kind="Table"]}[Data],
    #"Renamed Columns" = Table.RenameColumns(ShiftDemandHCAverageANACC_Table,{{"ShiftDurations.Duration", "ShiftDurations.ShiftDuration"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"ShiftPeriod", type text}, {"UnitShiftEffort", type number}, {"ShiftDurations.ShiftDuration", type number}, {"ShiftDemandHCAverage", type number}}),
    #"Filtered ROLE" = Table.SelectRows(#"Changed Type", each ([Role] = Role))
in
    #"Filtered ROLE";

shared #"IMPORT PermutationDimensions" = let
    Source = Excel.Workbook(File.Contents(FilePath&"\2. Calculations\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Shifts", type text}, 
 {"Period", Int64.Type}, {"RolesList", type text}}),
    #"Filtered ROLE" = Table.SelectRows(#"Changed Type", each ([RolesList] = Role))
in
    #"Filtered ROLE";

shared #"IMPORT Masterlist !!" = let
    Source = Excel.Workbook(File.Contents(FilePath&"\2. Calculations\StaffListMaster.xlsx"), null, true),
    Table_Masterlist_Table = Source{[Item="Table_Masterlist",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_Masterlist_Table,{{"Role", type text}}),
    #"Filtered ROLE" = Table.SelectRows(#"Changed Type", each ([Role] = Role))
in
    #"Filtered ROLE";

shared #"IMPORT ResDayShift !!" = let
    Source = Excel.Workbook(File.Contents(FilePath&"\2. Calculations\Capacity-ShiftAvailability.xlsx"), null, true),
    Table_ResDayShift_Table = Source{[Item="Table_ResDayShift",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_ResDayShift_Table,{{"Role", type text}, {"Name", type text},  {"Week", Int64.Type}, {"Day", type text}, {"Shift", type text}, {"EffectiveShiftHrs", type number}, {"Date", Int64.Type}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type",{{"Date", type date}}),
    #"Filtered ROLE" = Table.SelectRows(#"Changed Type1", each ([Role] = Role))
in
    #"Filtered ROLE";

shared #"IMPORT ResourceShiftAllocation - Role!!" = let
    Source = Excel.Workbook(File.Contents(FilePath&"\2. Calculations\AllocationByShiftAverage.xlsx"), null, true),
    ResourceShiftAllocation_Table = Source{[Item="ResourceShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResourceShiftAllocation_Table,{{"ShiftDate", type date}, {"ShiftPeriod", type text}, {"Name", type text}, {"Role", type text}, {"ResShiftEffort", type number}, {"ResShiftEffectiveRatio", type number}, {"ResShiftFTE", type number}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Role] = Role))
in
    #"Filtered Rows";

shared #"IMPORT MaxAvailability" = let
    Source = Excel.Workbook(File.Contents(FilePath &"\2. Calculations\Settings Data.xlsx"), null, true),
    MaxAvailability_Table = Source{[Item="MaxAvailability",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(MaxAvailability_Table,{{"MaxAvailability", Int64.Type}}),
    MaxAvailability1 = #"Changed Type"{0}[MaxAvailability]
in
    MaxAvailability1;

shared MaxShiftCluster = 5 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

[ Description = "Use to filter out shift creep into adjacent shifts" ]
shared NotShiftThreshold = 0.2 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared #"SET NWDPriority" = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("JYopEsAgEAS/Qq1ew+b28aGCpBA5QFFx/D/L4Lp7JgSyxHSmtz7XXRJFDiRa7G6OnKGDqhc2TqCjqhP2FjZhBM4YgEur/bC22nFT/GopFOMP", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [NWDPriority = _t, #"NWD Type" = _t])
in
    Source;

[ Description = "BUFFER" ]
shared #"PeriodShiftDay B" = let
    Source = #"IMPORT PermutationDimensions",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date", "Day", "Shifts", "Period", "RolesList"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

shared Resources = let
    Source = #"IMPORT Masterlist !!"
in
    Source;

shared #"ResourcePeriodTABLE-empty" = let
    Source = Resources,
    #"Added Custom" = Table.AddColumn(Source, "Periods", each #"PeriodShiftDay B"),
    #"Expanded Periods" = Table.ExpandTableColumn(#"Added Custom", "Periods", {"Period", "Shift", "Day"}, {"Period", "Shift", "Day"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Periods", each ([Period] <> "W")),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows",{{"Resource", Int64.Type}, {"Period", Int64.Type}})
in
    #"Changed Type";

[ Description = "BUFFER" ]
shared #"ResPeriodAvailabilityTABLE !!" = let
    Source = #"IMPORT ResDayShift !!",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Role", "Week", "Shift", "Date"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Date", type date}}),
    // Buffer
    BUFFER = Table.Buffer(#"Changed Type2"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Name"}, Resources, {"Name"}, "Resources", JoinKind.LeftOuter),
    #"Expanded Resources" = Table.ExpandTableColumn(#"Merged Queries", "Resources", {"Resource"}, {"Resource"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded Resources", {"Date", "Shift", "Role"}, #"PeriodShiftDay B", {"Date", "Shifts", "RolesList"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodShiftDay", {"Period"}, {"Period"}),
    // Make Availability dynamic
    #"Added AVAILABILITY TEMP" = Table.AddColumn(#"Expanded PeriodShiftDay", "Availability", each 1),
    #"Removed Other Columns1" = Table.SelectColumns(#"Added AVAILABILITY TEMP",{"Role", "Resource", "Period", "Availability"})
in
    #"Removed Other Columns1";

shared ResDayPeriodAvailabilityTABLE = let
    Source = #"ResPeriodAvailabilityTABLE !!",
    #"Merged Queries" = Table.NestedJoin(Source, {"Period", "Role"}, #"PeriodShiftDay B", {"Period", "RolesList"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Day"}, {"Day"})
in
    #"Expanded PeriodShiftDay";

shared ResDayAvailabilityTABLE = let
    Source = ResDayPeriodAvailabilityTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Resource", "Day"}, {{"DayAvailbility", each List.Sum([Availability]), type number}, {"DayShiftCount", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

shared ResPeriodAvailabilitySUM = let
    Source = List.Sum(#"ResPeriodAvailabilityTABLE !!"[Availability])
in
    Source;

shared ResPeriodAllocationCHECK = let
    Source = #"IMPORT ResourceShiftAllocation - Role!!",
    BUFFER = Table.Buffer(Source),
    #"Filtered DATEFROM" = Table.SelectRows(BUFFER, each ([ShiftDate] >= Date_From)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered DATEFROM",{"ResShiftEffort", "ResShiftEffectiveRatio"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Name", "Role"}, Resources, {"Name", "Role"}, "Resources", JoinKind.LeftOuter),
    #"Expanded Resources" = Table.ExpandTableColumn(#"Merged Queries", "Resources", {"Resource"}, {"Resource"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Resources",{{"ResShiftFTE", "Allocation"}})
in
    #"Renamed Columns";

[ Description = "BUFFER" ]
shared #"ResPeriodAllocation-" = let
    Source = ResPeriodAllocationCHECK,
    #"Filtered Rows NULL" = Table.SelectRows(Source, each ([Resource] <> null)),
    #"Added SHIFTNUMBER" = Table.AddColumn(#"Filtered Rows NULL", "ShiftNumber", each if [IntervalAssociatedShift] = "AM" then 1 else if [IntervalAssociatedShift] = "PM" then 2 else if [IntervalAssociatedShift] = "NIGHT" then 3 else null)
in
    #"Added SHIFTNUMBER";

shared #"ResPeriodAllocation-ShiftBias" = let
    Source = #"ResPeriodAllocation-",
    BUFFER = Table.Buffer(Source),
    #"Grouped Rows1" = Table.Group(BUFFER, {"IntervalAssociatedShift", "Resource", "ShiftDate"}, {{"Allocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows1", each ([Allocation] > NotShiftThreshold)),
    #"Added SHIFTINDEX" = Table.AddColumn(#"Filtered Rows", "ShiftIndex", each if [IntervalAssociatedShift] = "AM" then 1 else if [IntervalAssociatedShift] = "PM" then 2 else 3),
    #"Grouped SHIFTBIAS" = Table.Group(#"Added SHIFTINDEX", {"Resource"}, {{"ShiftBias", each List.Average([ShiftIndex]), type number}})
in
    #"Grouped SHIFTBIAS";

[ Description = "BUFFER" ]
shared ResPeriodAllocationTABLE = let
    Source = #"ResPeriodAllocation-",
    BUFFER = Table.Buffer(Source),
    #"Merged Queries1" = Table.NestedJoin(BUFFER, {"ShiftDate", "IntervalAssociatedShift"}, #"PeriodShiftDay B", {"Date", "Shifts"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Grouped ALLOCATION" = Table.Group(#"Expanded PeriodShiftDay", {"ShiftDate", "IntervalAssociatedShift", "Role", "Resource", "Period", "ShiftNumber"}, {{"Allocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered NOTFULLSHIFT" = Table.SelectRows(#"Grouped ALLOCATION", each ([Allocation] > NotShiftThreshold)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered NOTFULLSHIFT",{"Role", "Resource", "Period", "ShiftNumber", "Allocation"})
in
    #"Removed Other Columns";

shared ResDayAllocationTABLE = let
    Source = Table.NestedJoin(ResPeriodAllocationTABLE, {"Period"}, #"PeriodShiftDay B", {"Period"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay1" = Table.ExpandTableColumn(Source, "PeriodShiftDay", {"Day"}, {"Day"}),
    #"Group RESDAYALLOCATION" = Table.Group(#"Expanded PeriodShiftDay1", {"Resource", "Day"}, {{"ResDayAllocation", each List.Sum([Allocation]), type nullable number}, {"ShiftAverage", each List.Average([ShiftNumber]), type number}})
in
    #"Group RESDAYALLOCATION";

shared PeriodAllocationTABLE = let
    Source = ResPeriodAllocationTABLE,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Period", Int64.Type}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"Period", "Role"}, {{"A", each List.Sum([Allocation]), type nullable number}})
in
    #"Grouped Rows";

shared #"PeriodA-C' Neg(Surfit)" = let
    Source = #"PeriodC'",
    #"Merged Queries" = Table.NestedJoin(Source, {"Period"}, PeriodAllocationTABLE, {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodALLOCTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"A"}, {"A"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded PeriodALLOCTABLE",null,0,Replacer.ReplaceValue,{"A"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Replaced Value", "PeriodA-CNeg", each [A] - [#"PeriodC'"], type number),
    #"Reordered A-C'" = Table.ReorderColumns(#"Inserted Subtraction",{"Period", "A", "PeriodC'", "PeriodA-CNeg"}),
    #"Filtered Rows" = Table.SelectRows(#"Reordered A-C'", each ([#"PeriodA-CNeg"] < 0)),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared #"PeriodC'" = let
    Source = Table.NestedJoin(#"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, #"ResourcePeriodTABLE-empty", {"Resource", "Period"}, "ResourcePeriodTABLE-empty", JoinKind.LeftOuter),
    #"Expanded DAY" = Table.ExpandTableColumn(Source, "ResourcePeriodTABLE-empty", {"Day"}, {"Day"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded DAY", {"Resource", "Day"}, ResDayNWDTABLE, {"Resource", "Day"}, "ResDayNWDTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayNWDTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayNWDTABLE", {"PotentialAvailability"}, {"PotentialAvailability"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResDayNWDTABLE", each ([PotentialAvailability] = null)),
    #"Grouped C'" = Table.Group(#"Filtered Rows", {"Period"}, {{"PeriodC'", each List.Sum([Availability]), type number}}),
    Custom1 = Table.Buffer(#"Grouped C'")
in
    Custom1;

shared #"PeriodC'-D" = let
    Source = #"PeriodC'",
    Custom1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(Custom1, {"Period"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C'-D", each [#"PeriodC'"] - [D], type number),
    #"Inserted Division" = Table.AddColumn(#"Inserted Subtraction", "C'/D", each [#"PeriodC'"] / [D], type number)
in
    #"Inserted Division";

shared #"PeriodC'-DPos (ExcessPot)" = let
    Source = #"PeriodC'-D",
    #"Filtered Rows" = Table.SelectRows(Source, each [#"C'-D"] > 0)
in
    #"Filtered Rows";

shared #"ResC'" = let
    Source = #"C' ResSingleShiftDayTABLE",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"RosterAvailability", each List.Sum([Availability]), type number}})
in
    #"Grouped Rows";

shared #"ResAv-C'" = let
    Source = #"ResC'",
    #"Added AVAILCAP" = Table.AddColumn(Source, "RosterAvailabilityCAPPED", each if [RosterAvailability] > #"IMPORT MaxAvailability"
then #"IMPORT MaxAvailability"
else [RosterAvailability]),
    #"Inserted REDUCTION" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted REDUCTION",{{"AvailabilityReduction", "ResAv-C'"}})
in
    #"Renamed Columns";

[ Description = "BUFFER   Tag RP cell with R that need to reduced becuase over allocated" ]
shared #"ResPeriod(ExcessPot)SpareAvailabilityTABLE" = let
    Source = #"C' ResSingleShiftDayTABLE",
    #"BUFER 1" = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(#"BUFER 1", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource", "Period"}, ResPeriodWDTABLE, {"Resource", "Period"}, "ResPeriodWDTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodWDTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodWDTABLE", {"PotentialAvailability"}, {"PotentialAvailability"}),
    #"Filtered SPARE-ALLOCATIONNULL" = Table.SelectRows(#"Expanded ResPeriodWDTABLE", each ([Allocation] = null) and ([PotentialAvailability] = null)),
    #"Merged RESAv-C'" = Table.NestedJoin(#"Filtered SPARE-ALLOCATIONNULL", {"Resource"}, #"ResAv-C'", {"Resource"}, "ResRosterAvailability", JoinKind.LeftOuter),
    #"Expanded ResRosterAvailability" = Table.ExpandTableColumn(#"Merged RESAv-C'", "ResRosterAvailability", {"ResAv-C'"}, {"ResAv-C'"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResRosterAvailability", {"Period"}, #"PeriodC'-DPos (ExcessPot)", {"Period"}, "PeriodD-A", JoinKind.LeftOuter),
    #"Expanded SURPLUS-PERIODC'-D.POS" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodD-A", {"C'-D", "C'/D"}, {"C'-D", "C'/D"}),
    #"Merged Queries3" = Table.NestedJoin(#"Expanded SURPLUS-PERIODC'-D.POS", {"Resource", "Period"}, #"MultiDayPeriod-Remove", {"Resource", "AvailablePeriod"}, "MultiDayPeriod-Remove", JoinKind.LeftOuter),
    #"Expanded MULTIDAYPERIOD -REMOVE" = Table.ExpandTableColumn(#"Merged Queries3", "MultiDayPeriod-Remove", {"NoAllocationKeep"}, {"NoAllocationKeep"}),
    #"Filtered MULTISHIFTDAY-LOWERPRIORITY" = Table.SelectRows(#"Expanded MULTIDAYPERIOD -REMOVE", each ([NoAllocationKeep] <> false       )),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered MULTISHIFTDAY-LOWERPRIORITY",null,0,Replacer.ReplaceValue,{"C'-D"}),
    #"Added UNASSIGNEDAVAIL" = Table.AddColumn(#"Replaced Value", "UnassignedAvail", each if [#"C'-D"] > 0 and [#"ResAv-C'"] >0 
then [Availability]
else 0),
    #"Filtered UNASSIGNEDAVAIL" = Table.SelectRows(#"Added UNASSIGNEDAVAIL", each ([UnassignedAvail] <>0)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered UNASSIGNEDAVAIL",{"Availability", "C'-D", "C'/D", "NoAllocationKeep"}),
    #"Added COMBINE P+R" = Table.AddColumn(#"Removed Columns", "PRCell", each "P" & Text.From([Period]) & "-R" & Text.From([Resource]))
in
    #"Added COMBINE P+R";

[ Description = "BUFFER" ]
shared ResPeriodCapPrioritised = let
    Source = #"ResPeriod(ExcessPot)SpareAvailabilityTABLE",
    BUFFER = Table.Buffer(Source),
    #"Filtered Rows" = Table.SelectRows(BUFFER, each ([UnassignedAvail] > 0)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Period"}, #"PeriodC'-DPos (ExcessPot)", {"Period"}, "C-DPos (ExcessPot)", JoinKind.LeftOuter),
    #"Expanded zPeriodC-DPos (ExcessPot)" = Table.ExpandTableColumn(#"Merged Queries", "C-DPos (ExcessPot)", {"C'-D", "C'/D"}, {"C'-D", "C'/D"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded zPeriodC-DPos (ExcessPot)", {"Period"}, #"PeriodA-C' Neg(Surfit)", {"Period"}, "PeriodA-C", JoinKind.LeftOuter),
    #"Expanded PeriodA-C" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodA-C", {"PeriodA-CNeg"}, {"PeriodA-CNeg"}),
    #"Filtered A-C C-D" = Table.SelectRows(#"Expanded PeriodA-C", each [#"PeriodA-CNeg"] <=0 and [#"C'-D"] > 0),
    #"Reordered Columns1" = Table.ReorderColumns(#"Filtered A-C C-D",{"Resource", "Period", "PRCell", "UnassignedAvail", "ResAv-C'", "PeriodA-CNeg", "C'-D", "C'/D"}),
    #"Merged Queries2" = Table.NestedJoin(#"Reordered Columns1", {"Resource", "Period"}, ResPeriodShiftNWDTABLE, {"Resource", "Period"}, "ResPeriodAvailPrioritiesTABLE", JoinKind.LeftOuter),
    #"Expanded NWDPRIORITIES" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAvailPrioritiesTABLE", {"ClusterAllocation", "NWDPriority", "NWDType"}, {"ClusterAllocation", "NWDPriority", "NWDType"})
in
    #"Expanded NWDPRIORITIES";

[ Description = "BUFFER" ]
shared ResourcePeriodTABLEIndex = let
    Source = #"ResourcePeriodTABLE-empty",
    #"Added Index" = Table.AddIndexColumn(Source, "ResIndex", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

[ Description = "BUFFER" ]
shared #"AppliedAllocation+ActiveDays#" = let
    Source = Table.NestedJoin(ResourcePeriodTABLEIndex, {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(Source, "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Merged Queries" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource", "Period"}, #"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Added INVALIDALLOCATION" = Table.AddColumn(#"Expanded ResPeriodAvailabilityTABLE", "InvalidAllocation", each if [Allocation] <> null 
and [Availability] = null 
then true
else null),
    #"Renamed Columns" = Table.RenameColumns(#"Added INVALIDALLOCATION",{{"InvalidAllocation", "InValidAllocation"}})
in
    #"Renamed Columns";

shared ResDaysAllocated = let
    Source = #"AppliedAllocation+ActiveDays#",
    #"Replaced Value" = Table.ReplaceValue(Source,null,0,Replacer.ReplaceValue,{"Allocation"}),
    // Not true allocation.  1 used as indicator mostly allocated shift
    #"Added VALIDALLOCATION" = Table.AddColumn(#"Replaced Value", "ValidAllocation", each if [InValidAllocation] = true then 0 else if [Allocation] > ValidAllocation then 1 else 0),
    #"Grouped RESDAYALLOCATION" = Table.Group(#"Added VALIDALLOCATION", {"Resource", "Day"}, {{"ResDayAllocation", each List.Max([ValidAllocation]), type number}}),
    // Valid allocation only.  Not true allocation
    #"Added RESDAYALLOCATED" = Table.AddColumn(#"Grouped RESDAYALLOCATION", "ResDayAllocated", each if [ResDayAllocation] <> 0 then [ResDayAllocation] else null),
    #"Sorted Rows" = Table.Sort(#"Added RESDAYALLOCATED",{{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "ResIndex", 2, 1, Int64.Type)
in
    #"Added Index";

shared ResourcePeriodIndexLimit = let
    Source = ResDaysAllocated,
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"ResIndexStart", each List.Min([ResIndex]), type number}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
shared CumIndexes = let
    Source = ResDaysAllocated,
    // from ResDayAllocation
    #"Renamed ALLOCATION" = Table.RenameColumns(Source,{{"ResDayAllocation", "Allocation"}}),
    BUFFER = Table.Buffer(#"Renamed ALLOCATION"),
    #"Sorted BY RES INDEX" = Table.Sort(BUFFER,{{"ResIndex", Order.Ascending}}),
    #"Merged Queries" = Table.NestedJoin(#"Sorted BY RES INDEX", {"Resource"}, ResourcePeriodIndexLimit, {"Resource"}, "ResourcePeriodIndexLimit", JoinKind.LeftOuter),
    #"Expanded ResourcePeriodIndexLimit" = Table.ExpandTableColumn(#"Merged Queries", "ResourcePeriodIndexLimit", {"ResIndexStart"}, {"ResIndexStart"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded ResourcePeriodIndexLimit",null,0,Replacer.ReplaceValue,{"Allocation"}),
    #"Inserted FORWAR STEP" = Table.AddColumn(#"Replaced Value", "ForwardStep", each [ResIndex] - [ResIndexStart]+1, type number),
    LIST.BUFFER = List.Buffer(#"Inserted FORWAR STEP"[ResDayAllocation]),
    #"Added FORWARDCUM" = Table.AddColumn(#"Inserted FORWAR STEP", "CumulativeForward", each List.Sum(List.Range(#"Inserted FORWAR STEP"[ResDayAllocated], [ResIndexStart]-2
,[ForwardStep]))),
    #"Added Index" = Table.AddIndexColumn(#"Added FORWARDCUM", "Index", 0, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 1, 1, Int64.Type),
    #"Added Index2" = Table.AddIndexColumn(#"Added Index1", "Index.2", 2, 1, Int64.Type),
    #"Added Index3" = Table.AddIndexColumn(#"Added Index2", "Index.3", 3, 1, Int64.Type)
in
    #"Added Index3";

[ Description = "BUFFER" ]
shared AllocationCode = let
    Source = CumIndexes,
    BUFFER1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(BUFFER1, {"Index"}, BUFFER1, {"Index.3"}, "Reordered Columns1", JoinKind.LeftOuter),
    #"Expanded YESTERDAY3" = Table.ExpandTableColumn(#"Merged Queries", "Reordered Columns1", {"ResDayAllocated"}, {"Yesterday.Allocated3"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded YESTERDAY3", {"Index"}, #"Expanded YESTERDAY3", {"Index.2"}, "Reordered Columns", JoinKind.LeftOuter),
    #"Expanded YESTERDAY2" = Table.ExpandTableColumn(#"Merged Queries2", "Reordered Columns", {"Resource", "ResDayAllocated"}, {"Yesterday.Resource2", "Yesterday.Allocated2"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded YESTERDAY2", {"Index"}, #"Expanded YESTERDAY2", {"Index.1"}, "Added Index1", JoinKind.LeftOuter),
    #"Expanded YESTERDAY1" = Table.ExpandTableColumn(#"Merged Queries1", "Added Index1", {"Resource", "ResDayAllocated"}, {"Yesterday.Resource1", "Yesterday.Allocated1"}),
    #"Merged Queries3" = Table.NestedJoin(#"Expanded YESTERDAY1", {"Index.1"}, #"Expanded YESTERDAY1", {"Index"}, "Added Index2", JoinKind.LeftOuter),
    #"Expanded TOMORROW1" = Table.ExpandTableColumn(#"Merged Queries3", "Added Index2", {"Resource", "ResDayAllocated"}, {"Tomorrow.Resource1", "Tomorrow.Allocated1"}),
    #"Merged Queries4" = Table.NestedJoin(#"Expanded TOMORROW1", {"Index.2"}, #"Expanded TOMORROW1", {"Index"}, "Expanded TOMORROW1", JoinKind.LeftOuter),
    #"Expanded TOMORROW2" = Table.ExpandTableColumn(#"Merged Queries4", "Expanded TOMORROW1", {"Resource", "ResDayAllocated"}, {"Tomorrow.Resource2", "Tomorrow.Allocated2"}),
    #"Merged Queries5" = Table.NestedJoin(#"Expanded TOMORROW2", {"Index.3"}, #"Expanded TOMORROW2", {"Index"}, "Expanded Reordered Columns1", JoinKind.LeftOuter),
    #"Expanded TOMORROW3" = Table.ExpandTableColumn(#"Merged Queries5", "Expanded Reordered Columns1", {"ResDayAllocated"}, {"Tomorrow.Allocated3"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded TOMORROW3",null,0,Replacer.ReplaceValue,{"Yesterday.Resource1","Yesterday.Allocated3", "Yesterday.Allocated2", "Yesterday.Allocated1", "ResDayAllocated", "Tomorrow.Allocated1", "Tomorrow.Allocated2", "Tomorrow.Allocated3"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Replaced Value",{"Resource", "Day", "Allocation", "ResDayAllocated", "ResIndex", "ResIndexStart", "ForwardStep",  "CumulativeForward", "Yesterday.Allocated3", "Yesterday.Resource2", "Yesterday.Allocated2", "Yesterday.Resource1", "Yesterday.Allocated1", "Tomorrow.Resource1", "Tomorrow.Allocated1", "Tomorrow.Resource2", "Tomorrow.Allocated2", "Tomorrow.Allocated3"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Other Columns",{"Resource", "Day", "Allocation", "ResIndex", "ResIndexStart", "ForwardStep", "CumulativeForward", "Yesterday.Allocated3", "Yesterday.Allocated2", "Yesterday.Allocated1", "ResDayAllocated", "Tomorrow.Allocated1", "Tomorrow.Allocated2", "Tomorrow.Allocated3", "Yesterday.Resource2", "Yesterday.Resource1", "Tomorrow.Resource1", "Tomorrow.Resource2"}),
    CODE5 = Table.AddColumn(#"Reordered Columns", "Code", each Text.Combine({Text.From([Yesterday.Allocated2], "en-AU"), Text.From([Yesterday.Allocated1], "en-AU"), Text.From([ResDayAllocated], "en-AU"), Text.From([Tomorrow.Allocated1], "en-AU"), Text.From([Tomorrow.Allocated2], "en-AU")}, ""), type text),
    #"Added RESOURCEEND" = Table.AddColumn(CODE5, "ResourceStartEnd", each if [Resource] <> [Tomorrow.Resource1] then "ResourceEnd1" else if [Resource] <> [Yesterday.Resource1] then "ResourceStart1" else if [Resource] <> [Tomorrow.Resource2] then "ResourceEnd2" else if [Resource] <> [Yesterday.Resource2] then "ResourceStart2" else null),
    #"Removed Other Columns1" = Table.SelectColumns(#"Added RESOURCEEND",{"Resource", "Day", "Allocation", "ResIndex", "ResDayAllocated", "Tomorrow.Resource1", "Code", "ResourceStartEnd"}),
    #"Added CLUSTERPOINTS" = Table.AddColumn(#"Removed Other Columns1", "ClusterPoints", each 
        if (                    
            //general start
            (  [Code] = "00111" 
                or [Code] ="00101" 
                or [Code] ="00110" 
            )
            //Resouce start
            or (
                (   [ResourceStartEnd] = "ResourceStart1"
                    or
                    [ResourceStartEnd] = "ResourceStart2"
                )
                and (
                    [Code] = "11111"
                    or 
                    [Code] = "10111"

                    )
                )           
           )     
        then 
        "Allocated-ClusterStart"

        else 
            if 
            //General end
            [Code] = "11100" 
            or [Code] ="10100" 
            or [Code] ="01100"
            //Change of resource
            or  (
                [Tomorrow.Resource1] <> [Resource] 
                and [Allocation] <> 0
                )
           
        then "Allocated-ClusterEnd" 

        else 
            if 
            //General end
            [Code] = "00100" 
            then "Allocated-SingleStart/End"        
        else null),
    #"Reordered Columns2" = Table.ReorderColumns(#"Added CLUSTERPOINTS",{"Resource", "Day", "Allocation", "ResIndex", "ResDayAllocated", "Code", "ResourceStartEnd", "ClusterPoints"}),
    BUFFER = Table.Buffer(#"Reordered Columns2")
in
    BUFFER;

shared ClusterIDStart = let
    Source = #"AllocationCode",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Day", "ClusterPoints"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each [ClusterPoints] = "Allocated-ClusterStart" or [ClusterPoints] = "Allocated-ClusterSINGLE" or [ClusterPoints] = "Allocated-SingleStart/End"),
    #"Added Index1" = Table.AddIndexColumn(#"Filtered Rows", "ClusterIndex", 1, 1, Int64.Type)
in
    #"Added Index1";

shared ClusterIDEnd = let
    Source = #"AllocationCode",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Day", "ClusterPoints"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each [ClusterPoints] = "Allocated-ClusterEnd" or [ClusterPoints] = "Allocated-ClusterSINGLE"),
    #"Added Index1" = Table.AddIndexColumn(#"Filtered Rows", "ClusterIndex", 1, 1, Int64.Type)
in
    #"Added Index1";

shared #"ClusterIDStart+End" = let
    Source = Table.Combine({ClusterIDEnd, ClusterIDStart}),
    #"Removed Duplicates" = Table.Distinct(Source)
in
    #"Removed Duplicates";

[ Description = "BUFFER" ]
shared Clusters = let
    Source = #"AllocationCode",
    #"BUFFER 1" = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(#"BUFFER 1", {"Resource", "Day", "ClusterPoints"}, #"ClusterIDStart+End", {"Resource", "Day", "ClusterPoints"}, "ClusterIDStart+End", JoinKind.LeftOuter),
    #"Expanded ClusterIDStart+End" = Table.ExpandTableColumn(#"Merged Queries", "ClusterIDStart+End", {"ClusterIndex"}, {"ClusterIndex"}),
    #"Filled Down" = Table.FillDown(#"Expanded ClusterIDStart+End",{"ClusterIndex"}),
    #"Filled Up" = Table.FillUp(#"Filled Down",{"ClusterIndex"}),
    #"Renamed Columns" = Table.RenameColumns(#"Filled Up",{{"ClusterIndex", "ClusterIndexX"}}),
    #"Added NO ALLOCATION NULL" = Table.AddColumn(#"Renamed Columns", "ClusterIndex", each if [Allocation] = 0 then null else [ClusterIndexX]),
    #"Removed Columns" = Table.RemoveColumns(#"Added NO ALLOCATION NULL",{"ClusterIndexX", "ClusterPoints"}),

// Now we add the "ClusterNumber" column based on the previous row's value
AddPreviousClusterIndex = Table.AddColumn(#"Removed Columns", "ClusterNumber", each 
    let
        // Determine the steps needed based on the Code value
        RowSteps = if [Code] = "00000" then "X" else
                   if List.Contains({"11000", "11011", "01010", "01001", "01000", "11001","11010"}, [Code]) then -1 else
                   if List.Contains({"10000", "10001", "10010"}, [Code]) then -2 else
                   if List.Contains({"00011", "00010"}, [Code]) then 1 else
                   if List.Contains({"00001", "10011"}, [Code]) then 2 else
                   0,

        // Calculate the previous index
        PrevIndex = if RowSteps = "X" then null else [ResIndex] + RowSteps,

        // Retrieve the previous row based on the calculated index
        PrevRowT = if PrevIndex = null then null else Table.SelectRows(#"Removed Columns", each [ResIndex] = PrevIndex),

        // Extract the ClusterIndex from the previous row
        PrevRow = if PrevRowT = null or Table.IsEmpty(PrevRowT) then "X"  else PrevRowT{0}[ClusterIndex]
    in
        PrevRow),
    BUFFER = Table.Buffer(AddPreviousClusterIndex),
    #"Sorted Rows" = Table.Sort(BUFFER,{{"Resource", Order.Ascending}})
in
    #"Sorted Rows";

shared ClusterAllocation = let
    Source = Clusters,
    #"Grouped CLUSTERSIZE" = Table.Group(Source, {"ClusterIndex"}, {{"ClusterAllocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped CLUSTERSIZE", each ([ClusterIndex] <> null))
in
    #"Filtered Rows";

[ Description = "BUFFER #(lf)S2, P2 dev debt #(lf)Opportunity to discriminate with small Cluster Allocation#(lf)i.e. change the cluster it is associated with" ]
shared #"WD Type B" = let
    Source = Clusters,
    Custom1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(Custom1, {"ClusterNumber"}, ClusterAllocation, {"ClusterIndex"}, "ClusterAllocation", JoinKind.LeftOuter),
    #"Expanded ClusterAllocation" = Table.ExpandTableColumn(#"Merged Queries", "ClusterAllocation", {"ClusterAllocation"}, {"ClusterAllocation"}),
    #"BUFFER LIST" = List.Buffer(#"Expanded ClusterAllocation"[Code]),
    // S2, P2 dev debt 
    // Opportunity to discriminate with small Cluster Allocation
    #"Added NWD TYPE" = Table.AddColumn(#"Expanded ClusterAllocation", "NWDType", each if [ResDayAllocated] = 1  then "Allocated" 

        else if [Code] = "00000" then "Reducable"

    //S1
        else if [Code] = "01000" 
            and [ResourceStartEnd] <> "ResourceStart1"  
        then "S1"

        else if [Code] = "11000" 
            and [ResourceStartEnd]<>"ResourceStart1" 
        then "S1"


     //S2       
        else if [Code] = "10000" 
            and [ResourceStartEnd] <> "ResourceStart2"  
            and [ResourceStartEnd] <> "ResourceStart1"         
        then "S2"

    //S2 P2

        else if [Code] = "10001" 
            and [ResourceStartEnd]<>"ResourceStart2" 
            and [ResourceStartEnd]<>"ResourceEnd2"         
        then "S2, P2"


    //P2
        else if (
                    [Code] = "00001" 
                    and [ResourceStartEnd] <> "ResourceEnd2" 
                    and [ResourceStartEnd] <> "ResourceEnd1"
                )
                or 
                (
                    [Code] = "11001" 
                    and [ResourceStartEnd] = "ResourceEnd2" 
                )
        then "P2"


    //P1
     
        else if (
                    ([Code] = "00010" or [Code] = "00011" )
                    and [ResourceStartEnd] <> "ResourceEnd1"    
                )  
                or 
                (
                    [Code] = "11011"
                    and [ResourceStartEnd] = "ResourceEnd1"
                )
        then "P1"
        

    //2d Off

        else if [Code] = "01001" 
            then
                if [ResourceStartEnd]="ResourceStart1"  
                then "P2"
            else "P2,S1"

        else if [Code] = "10010" 
            then 
                if [ResourceStartEnd]="ResourceEnd1"  
                then  "S2"
            else "P2,S1"
        
        else if [Code] = "10011"
            then  
                if [ResourceStartEnd]="ResourceStart1"
                then  "P1"
            else "P2,S1"

        else if [Code] = "11001" 
            then 
                if [ResourceStartEnd] = "ResourceEnd1"
                then "S1"
            else "P2"
            

    //Conditional

        
        else if [Code] = "01010" 
               and ("Yesterday.Alloction3" = 1 
               or "Tomorrow.Alloction3" = 1) 
               then "1D Off Conditional"

        else if [Code] = "01011" 
               and ("Yesterday.Alloction3" = 0 
               or "Tomorrow.Alloction3" = 0) 
               then "1D Off Conditional"


        
    //1D Off

        else if [Code] = "11010" 
            and [ResourceStartEnd] <> "ResourceEnd1" 
            and [ResourceStartEnd] <> "ResourceStart1"
        then "1D Off"
     
        else if [Code] = "11011" 
            and [ResourceStartEnd] <> "ResourceEnd1" 
            and [ResourceStartEnd] <> "ResourceStart1"
        then "1D Off"
              
        else "Reducable"),
    #"Removed Other Columns" = Table.SelectColumns(#"Added NWD TYPE",{"Resource", "Day",  "ClusterNumber", "ClusterAllocation", "NWDType"}),
    Custom2 = #"Removed Other Columns"
in
    Custom2;

shared ResDayPotentialAvailabilityTABLE = let
    Source = #"WD Type B",
  


    #"POTENTIAL ALLOCATION" = Table.AddColumn(Source, "PotentialAvailability", each if    [NWDType] = "Allocated" then "U"
    
    else if [NWDType] = "Reducable" then "R"

    else if [ClusterAllocation] >= MaxShiftCluster then "R"
    
    else if [ClusterAllocation] = (MaxShiftCluster -1)
        and ([NWDType] = "1D Off"   or  [NWDType] <> "S1" or [NWDType] <> "P1") 
        then "R" 
    
    else if [ClusterAllocation] <= (MaxShiftCluster -2)
        and  ([NWDType] = "1D Off"   
        or  [NWDType] <> "S1" 
        or [NWDType] <> "P1" 
        or [NWDType] <> "S2"  
        or [NWDType] <> "S2,P2" 
        or [NWDType] <> "P1,S2"
        or [NWDType] <> "P2,S1")
        then "R"
      

   
    else "U"),
    BUFFER = Table.Buffer(#"POTENTIAL ALLOCATION")
in
    BUFFER;

[ Description = "BUFFER" ]
shared ResPeriodShiftNWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "R")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"PotentialAvailability"}),
    BUFFER = Table.Buffer(#"Removed Columns"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"NWDType"}, #"SET NWDPriority", {"NWD Type"}, "NWDPrioritiesTABLE", JoinKind.LeftOuter),
    #"Expanded NWDPrioritiesTABLE" = Table.ExpandTableColumn(#"Merged Queries", "NWDPrioritiesTABLE", {"NWDPriority"}, {"NWDPriority"}),
    #"Sorted Rows" = Table.Sort(#"Expanded NWDPrioritiesTABLE",{{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows", {"Day"}, #"PeriodShiftDay B", {"Day"}, "Period.1", JoinKind.LeftOuter),
    #"Expanded PERIODSHIFTS" = Table.ExpandTableColumn(#"Merged Queries1", "Period.1", {"Period", "Shifts"}, {"Period", "Shifts"})
in
    #"Expanded PERIODSHIFTS";

shared ResDayNWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "R")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Resource", "Day", "PotentialAvailability"})
in
    #"Removed Other Columns";

[ Description = "BUFFER" ]
shared ResPeriodWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "U")),
    BUFFER = Table.Buffer(#"Filtered Rows"),
    #"Removed Other Columns" = Table.SelectColumns(BUFFER,{"Resource", "Day", "PotentialAvailability"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Other Columns", {"Day"}, #"PeriodShiftDay B", {"Day"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodShiftDay", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE (2)", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE (2)", {"Allocation"}, {"Allocation"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded ResPeriodAllocationTABLE", {"Resource", "Period"}, #"C' ResSingleShiftDayTABLE", {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded C'" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Sorted Rows" = Table.Sort(#"Expanded C'",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    #"Replaced Value" = Table.ReplaceValue(#"Sorted Rows",null,0,Replacer.ReplaceValue,{"Allocation", "Availability"}),
    #"Added ROSTEREDPERIODSTATUS" = Table.AddColumn(#"Replaced Value", "RosteredPeriodStatus", each if ([Availability] = null or [Availability] = 0) 
and 
[Allocation] > 0  
then "UnavailableAllocatedPeriod"

else if [Allocation] = 0 then "RosteredDay" 
else if [Allocation] <> 0 then "AllocatedPeriod" 
else null),

    #"Removed Columns" = Table.RemoveColumns(#"Added ROSTEREDPERIODSTATUS",{"Allocation"})
in
    #"Removed Columns";

shared UnavailableAllocated = let
    Source = Table.NestedJoin(#"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.FullOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(Source, "ResPeriodAllocationTABLE", {"Resource", "Period", "Allocation"}, {"Resource.1", "Period.1", "Allocation"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded ResPeriodAllocationTABLE", each ([Availability] = null or [Availability] = 0 ) and ([Allocation] <> null )),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Resource", "Period", "Availability"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Resource.1", "Resource"}, {"Period.1", "Period"}}),
    #"Added UNAVAILABLEALLOCATIONPERIOD" = Table.AddColumn(#"Renamed Columns", "UnavailableAllocated", each "UnavailableAllocationPeriod")
in
    #"Added UNAVAILABLEALLOCATIONPERIOD";

shared #"SingleShiftDay-initial" = let
    Source = ResDayAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([DayShiftCount] = 1))
in
    #"Filtered Rows";

[ Description = "BUFFER" ]
shared #"MultiDayPeriod-Available+Allocation" = let
    Source = ResDayAvailabilityTABLE,
    #"Filtered MULTIDAY" = Table.SelectRows(Source, each ([DayShiftCount] <> 1)),
    BUFFER = Table.Buffer(#"Filtered MULTIDAY"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Resource", "Day"}, ResDayPeriodAvailabilityTABLE, {"Resource", "Day"}, "ResDayPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayPeriodAvailabilityTABLE", {"Period"}, {"AvailablePeriod"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResDayPeriodAvailabilityTABLE", {"Resource", "AvailablePeriod"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded ResPeriodAllocationTABLE",{"DayAvailbility", "DayShiftCount"})
in
    #"Removed Columns";

shared MultiShiftAvilabilityDay = let
    Source = #"MultiDayPeriod-Available+Allocation",
    #"Removed MULTIPERIODALLOCATION" = Table.RemoveColumns(Source,{"Allocation"}),
    BUFFER = Table.Buffer(#"Removed MULTIPERIODALLOCATION"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Resource", "Day"}, ResDayAllocationTABLE, {"Resource", "Day"}, "ResDayAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayAllocationTABLE", {"ResDayAllocation"}, {"ResDayAllocation"}),
    #"Filtered UNMACTHED PERIODS REMOVED" = Table.SelectRows(#"Expanded ResDayAllocationTABLE", each ([ResDayAllocation] = null))
in
    #"Filtered UNMACTHED PERIODS REMOVED";

shared #"MultishiftAvailDayPeriod-Index" = let
    Source = MultiShiftAvilabilityDay,
    #"BUFFER 1" = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(#"BUFFER 1", {"Resource", "AvailablePeriod"}, ResDayPeriodAvailabilityTABLE, {"Resource", "Period"}, "ResDayPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResDayPeriodAvailabilityTABLE", {"AvailablePeriod"}, PeriodCapacityTABLE, {"Period"}, "PeriodCapacityTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodCapacityTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodCapacityTABLE", {"Capacity"}, {"Capacity"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded PeriodCapacityTABLE", {"AvailablePeriod"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted C/D" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C/D", each [Capacity] / [D], type number),
    #"Merged Queries3" = Table.NestedJoin(#"Inserted C/D", {"Resource"}, #"ResPeriodAllocation-ShiftBias", {"Resource"}, "ResPeriodAllocation-ShiftBias", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocation-ShiftBias" = Table.ExpandTableColumn(#"Merged Queries3", "ResPeriodAllocation-ShiftBias", {"ShiftBias"}, {"ShiftBias"}),
    #"Sorted RES, PERIOD C/D" = Table.Sort(#"Expanded ResPeriodAllocation-ShiftBias",{{"Resource", Order.Ascending}, {"C/D", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted RES, PERIOD C/D", "Index", 1, 1, Int64.Type),
    #"Removed Other Columns" = Table.SelectColumns(#"Added Index",{"Resource", "Day", "AvailablePeriod", "C/D", "ShiftBias", "Index", "Role"})
in
    #"Removed Other Columns";

shared #"MultiShiftAvailDayPeriod -0 Allocation" = let
    Source = #"MultishiftAvailDayPeriod-Index",
    BUFFER1 = Table.Buffer(Source),
    #"Grouped RESDAYS" = Table.Group(BUFFER1, {"Resource", "Day"}, {{"ResGroup", each _, type table [Role=nullable text, Resource=nullable number, Day=number, AvailablePeriod=nullable number, ResDayAllocation=nullable number, Availability=number, Capacity=nullable number, D=nullable number, #"C/D"=number, ShiftBias=nullable number]}}),
    #"Added DAY INDEX" = Table.AddColumn(#"Grouped RESDAYS", "Group DAY INDEX", each Table.AddIndexColumn([ResGroup],"ResC/DPriority",1)),
    #"Expanded DAYINDEX" = Table.ExpandTableColumn(#"Added DAY INDEX", "Group DAY INDEX", {"AvailablePeriod", "C/D", "ResC/DPriority", "ShiftBias"}, {"AvailablePeriod", "C/D", "ResC/DPriority", "ShiftBias"}),
    #"Added NO ALLOCATION KEEP" = Table.AddColumn(#"Expanded DAYINDEX", "NoAllocationKeep", each if [#"ResC/DPriority"] = 1 then true else false),
    #"Changed Type" = Table.TransformColumnTypes(#"Added NO ALLOCATION KEEP",{{"NoAllocationKeep", type logical}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Changed Type",{"ResGroup", "ResC/DPriority"})
in
    #"Removed Columns3";

shared #"SingleShiftDayAvail+MultiAllocation B" = let
    Source = #"SingleShiftDay-initial",
    BUFFER = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Day"}, #"PeriodShiftDay B", {"Day"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodShiftDay", {"Resource", "Period", "Role"}, ResPeriodAllocationTABLE, {"Resource", "Period", "Role"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodAllocationTABLE", {"Allocation"}, {"Allocation"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded ResPeriodAllocationTABLE",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    #"Filtered ALLOCATED PERIODS" = Table.SelectRows(#"Sorted Rows1", each ([Allocation] <> null )),
    #"Merged Queries2" = Table.NestedJoin(#"Filtered ALLOCATED PERIODS", {"Resource", "Period"}, #"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Filtered NULL AVAILABILITY" = Table.SelectRows(#"Expanded ResPeriodAvailabilityTABLE", each ([Availability] <> null )),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered NULL AVAILABILITY",{"Period", "Resource"})
in
    #"Removed Other Columns";

shared SingleShiftDayPeriodMatch = let
    Source = DayShiftMatchSPLIT,
    #"Filtered Rows" = Table.SelectRows(Source, each ([MatchCount] = "SingleShiftMatch")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"DayMatches", "MatchCount"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Day", "Resource"}, ResDayPeriodAvailabilityTABLE, {"Day", "Resource"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded ResDayPeriodAvailability" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period", "Availability"}, {"Period", "Availability"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Expanded ResDayPeriodAvailability",{"Day"}),
    #"Added MATCHPERIOD?" = Table.AddColumn(#"Removed Columns1", "MatchedPeriood", each if [MatchPeriod] = [Period] then true else false, type logical),
    #"Filtered Rows1" = Table.SelectRows(#"Added MATCHPERIOD?", each ([MatchedPeriood] = true)),
    #"Removed Columns2" = Table.RemoveColumns(#"Filtered Rows1",{"MatchedPeriood", "MatchPeriod", "Availability"})
in
    #"Removed Columns2";

[ Description = "BUFFER" ]
shared #"C' ResSingleShiftDayTABLE" = let
    Source = Table.Combine({MultiDayPeriodMatchTABLE, SingleShiftDayPeriods}),
    #"Appended MULTIDAYPERIOD-REDUCIBLE" = Table.Combine({Source, #"MultiDayPeriod-Reducibile"}),
    #"Removed Duplicates" = Table.Distinct(#"Appended MULTIDAYPERIOD-REDUCIBLE"),
    #"Merged Queries" = Table.NestedJoin(#"Removed Duplicates", {"Resource", "Period"}, #"ResPeriodAvailabilityTABLE !!", {"Resource", "Period"}, "ResPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"BUFFER Expanded ResPeriodAvailabilityTABLE" = Table.Buffer(Table.ExpandTableColumn(#"Merged Queries", "ResPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}))
in
    #"BUFFER Expanded ResPeriodAvailabilityTABLE";

shared #"C' SUM" = let
    Source = List.Sum(#"C' ResSingleShiftDayTABLE"[Availability])
in
    Source;

shared #"MultiDayPeriod-AvailableAllocatedMatches" = let
    Source = #"MultiDayPeriod-Available+Allocation",
  #"Filtered NULL ALLOCATION PERIOD" = Table.SelectRows(Source, each [Allocation] <> null)
in
  #"Filtered NULL ALLOCATION PERIOD"
  //Source
;

[ Description = "BUFFER" ]
shared DayShiftMatchSPLIT = let
    Source = #"MultiDayPeriod-AvailableAllocatedMatches",
    #"Grouped DAYMATCHES" = Table.Group(Source, {"Resource", "Day"}, {{"DayMatches", each Table.RowCount(_), Int64.Type}, {"MatchPeriod", each List.Average([AvailablePeriod]), type nullable number}}),
    #"SPLIT BUFFER" = Table.Buffer(Table.AddColumn(#"Grouped DAYMATCHES", "MatchCount", each if [DayMatches] = 1 then "SingleShiftMatch" else "MultipleShiftMatch"))
in
    #"SPLIT BUFFER";

shared MultiShiftDayMatches = let
    Source = DayShiftMatchSPLIT,
    #"Filtered Rows" = Table.SelectRows(Source, each ([MatchCount] = "MultipleShiftMatch")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"MatchCount", "MatchPeriod"})
in
    #"Removed Columns";

shared MultiShiftDayPeriodMatches = let
    Source = MultiShiftDayMatches,
    BUFFER = Table.Buffer(Source),
    #"Merged Queries1" = Table.NestedJoin(BUFFER, {"Resource", "Day"}, #"MultiDayPeriod-AvailableAllocatedMatches", {"Resource", "Day"}, "MultiDayPeriod-AvailableAllocatedMatches", JoinKind.LeftOuter),
    #"Expanded MultiDayPeriod-AvailableAllocatedMatches" = Table.ExpandTableColumn(#"Merged Queries1", "MultiDayPeriod-AvailableAllocatedMatches", {"AvailablePeriod"}, {"AvailablePeriod"})
in
    #"Expanded MultiDayPeriod-AvailableAllocatedMatches";

shared #"MultiPeriodPriority B" = let
    Source = MultiShiftDayPeriodMatches[[AvailablePeriod]],
    #"Removed Duplicates" = Table.Distinct(Source),
    BUFFER = Table.Buffer(#"Removed Duplicates"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"AvailablePeriod"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded PeriodDemandTABLE", {"AvailablePeriod"}, PeriodAllocationTABLE, {"Period"}, "PeriodAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodAllocationTABLE", {"A"}, {"A"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded PeriodAllocationTABLE", {"AvailablePeriod"}, PeriodCapacityTABLE, {"Period"}, "PeriodCapacityTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodCapacityTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "PeriodCapacityTABLE", {"Capacity"}, {"Capacity"}),
    #"Inserted Division" = Table.AddColumn(#"Expanded PeriodCapacityTABLE", "C/D", each [Capacity] / [D], type number),
    #"Inserted Division1" = Table.AddColumn(#"Inserted Division", "A/D", each [A] / [D], type number),
    #"Sorted C/D A/D" = Table.Sort(#"Inserted Division1",{{"C/D", Order.Ascending}, {"A/D", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted C/D A/D", "PeriodPriority", 1, 1, Int64.Type),
    #"Removed Columns" = Table.RemoveColumns(#"Added Index",{"D", "A", "Capacity", "C/D", "A/D"})
in
    #"Removed Columns";

shared MultiResPeriodPriority = let
    Source = Table.NestedJoin(MultiShiftDayPeriodMatches, {"AvailablePeriod"}, #"MultiPeriodPriority B", {"AvailablePeriod"}, "MultiPeriodPriority", JoinKind.LeftOuter),
    #"Expanded MultiPeriodPriority" = Table.ExpandTableColumn(Source, "MultiPeriodPriority", {"PeriodPriority"}, {"PeriodPriority"})
in
    #"Expanded MultiPeriodPriority";

shared #"MultiResPeriodPriority-Highest" = let
    Source = MultiResPeriodPriority,
    #"Grouped Rows" = Table.Group(Source, {"Resource", "Day"}, {{"HighestPriority", each List.Max([PeriodPriority]), type nullable number}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
shared MultiDayPeriodMatchTABLE = let
    Source = Table.NestedJoin(MultiResPeriodPriority, {"Resource", "Day"}, #"MultiResPeriodPriority-Highest", {"Resource", "Day"}, "MultiResPriodPriority-Highest", JoinKind.LeftOuter),
    #"Expanded MultiResPriodPriority-Highest" = Table.ExpandTableColumn(Source, "MultiResPriodPriority-Highest", {"HighestPriority"}, {"HighestPriority"}),
    #"Added Custom" = Table.AddColumn(#"Expanded MultiResPriodPriority-Highest", "Highest?", each [PeriodPriority]=[HighestPriority]),
    #"Filtered HIGHEST PRIORITY" = Table.SelectRows(#"Added Custom", each ([#"Highest?"] = true)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered HIGHEST PRIORITY",{"DayMatches", "PeriodPriority", "HighestPriority", "Highest?"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailablePeriod", "Period"}}),
    #"Removed Columns1" = Table.Buffer(Table.RemoveColumns(#"Renamed Columns",{"Day"}))
in
    #"Removed Columns1";

shared SingleShiftDayPeriods = let
    // SingleShiftDayAvail+Multi + SingleDayShiftMatchTABLE
    Source = Table.Combine({SingleShiftDayPeriodMatch, #"SingleShiftDayAvail+MultiAllocation B"}),
    #"Sorted Rows" = Table.Sort(Source,{{"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows";

[ Description = "BUFFER" ]
shared #"MultiDayPeriod-Reducibile" = let
    Source = #"MultiShiftAvailDayPeriod -0 Allocation",
    BUFFER = Table.Buffer(Source),
    #"Filtered Rows" = Table.SelectRows(BUFFER, each ([NoAllocationKeep] = true)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Day", "NoAllocationKeep", "C/D"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailablePeriod", "Period"}})
in
    #"Renamed Columns";

shared #"MultiDayPeriod-Remove" = let
    Source = #"MultiShiftAvailDayPeriod -0 Allocation",
    #"Filtered Rows" = Table.SelectRows(Source, each [NoAllocationKeep] = false),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{ "Day"})
in
    #"Removed Columns";

shared PeriodAllocated = let
    Source = ResPeriodAllocationTABLE,
    #"Merged Queries" = Table.NestedJoin(Source, {"Period"}, #"PeriodShiftDay B", {"Period"}, "PeriodShiftDay", JoinKind.LeftOuter)
in
    #"Merged Queries";

shared PeriodCapacityTABLE = let
    Source = #"ResPeriodAvailabilityTABLE !!",
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"Capacity", each List.Sum([Availability]), type number}})
in
    #"Grouped Rows";

shared ResPeriodAllocationSUM = let
    Source = List.Sum(PeriodAllocationTABLE[A])
in
    Source;

[ Description = "BUFFER" ]
shared #"Demand Prepare" = let
    Source = #"IMPORT ShiftUnitDemandHRS !!",
    #"Renamed Columns1" = Table.RenameColumns(Source,{{"ShiftPeriod", "Shift"}, {"ShiftDemandHCAverage", "DemandFTE"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns1",{"UnitShiftEffort", "ShiftDurations.ShiftDuration"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Removed Columns",{"Shift", "Role", "Facility", "DemandFTE", "Date"}),
    BUFFER = Table.Buffer(#"Removed Other Columns"),
    #"Merged Queries" = Table.NestedJoin(BUFFER, {"Date", "Shift"}, #"PeriodShiftDay B", {"Date", "Shifts"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded PeriodShiftDay",{{"DemandFTE", "D"}})
in
    #"Renamed Columns";

shared #"PeriodDemandTABLE !!" = let
    Source = #"Demand Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Shift", "Role", "D", "Period"})
in
    #"Removed Other Columns";

shared PeriodDemandSUM = let
    Source = List.Sum(#"PeriodDemandTABLE !!"[D])
in
    Source;

shared Date_From = let
    Source = Excel.Workbook(File.Contents(FilePath&"\2. Calculations\Settings Data.xlsx"), null, true),
    DateFrom_Table = Source{[Item="DateFrom",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(DateFrom_Table,{{"DateFrom", type date}}),
    DateFrom = #"Changed Type1"{0}[DateFrom]
in
    DateFrom;

shared ValidAllocation = 0.7 meta [IsParameterQuery=true, Type="Number", IsParameterQueryRequired=true];

shared #"PeriodC'SUM" = let
    Source = #"PeriodC'",
    #"PeriodC'1" = Source[#"PeriodC'"],
    #"Calculated Sum" = List.Sum(#"PeriodC'1")
in
    #"Calculated Sum";

shared RolePathTABLE = let
    // PART A - Define FilePathUrl and Replaced Value
    FilePathUrl = 
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        #"Renamed Columns" = Table.RenameColumns(Source, {{"Column1", "FilePath"}}),
        ReplacedValue = Table.ReplaceValue(#"Renamed Columns", "/", "\", Replacer.ReplaceText, {"FilePath"}),
        BufferedTable = Table.Buffer(ReplacedValue) // Buffer the table for better performance
    in 
        BufferedTable,

    // PART B - RootPath
    IMPORTRootPath =
    let
        // Step 1: Import CentriSyncPaths
        CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
        CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
        CentriSyncPaths_ChangedType = Table.TransformColumnTypes(CentriSyncPaths_Table, {{"User", type text}, {"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),

        // Step 2: UrlSite - Use buffered FilePathUrl
        UrlSite_ExtractedTextAfterDelimiter = Table.TransformColumns(FilePathUrl, {{"FilePath", each Text.AfterDelimiter(_, "sites\"), type text}}),
        UrlSite_ExtractedTextBeforeDelimiter = Table.TransformColumns(UrlSite_ExtractedTextAfterDelimiter, {{"FilePath", each Text.BeforeDelimiter(_, "\"), type text}}),
        RenamedColumns1 = Table.RenameColumns(UrlSite_ExtractedTextBeforeDelimiter, {{"FilePath", "Site"}}),

        // Step 3: Prefix
        Prefix_NestedJoin = Table.NestedJoin(RenamedColumns1, {"Site"}, CentriSyncPaths_ChangedType, {"Site"}, "CentriSyncPaths", JoinKind.LeftOuter),
        Prefix_Expanded = Table.ExpandTableColumn(Prefix_NestedJoin, "CentriSyncPaths", {"SyncedFolderRootPath"}, {"Prefix"}),
        Prefix = Prefix_Expanded{0}[Prefix],

        // Step 4: Core
        Core_ExtractedTextAfterDelimiter = Table.TransformColumns(FilePathUrl, {{"FilePath", each Text.AfterDelimiter(_, "Shared Documents\"), type text}}),
        Core_ExtractedTextBeforeDelimiter = Table.TransformColumns(Core_ExtractedTextAfterDelimiter, {{"FilePath", each Text.BeforeDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
        Core = Core_ExtractedTextBeforeDelimiter{0}[FilePath],

        // Step 5: FilePath
        FilePath = Prefix & "\" & Core,
        ConvertedToTable = #table(1, {{FilePath}}),
        RenamedColumns = Table.RenameColumns(ConvertedToTable, {{"Column1", "FilePath"}})
    in  
        RenamedColumns,

    // PART C - Dimensions
    // Extract UserName directly
    UserName = 
    let
        Source = IMPORTRootPath,
        Extracted = Text.BeforeDelimiter(Text.AfterDelimiter(Source[FilePath]{0}, "\", 1), "\")
    in
        Extracted,

    // Extract Client directly
    Client = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{3, RelativePosition.FromEnd}), // Extract the unit before the last delimiter
                    "\")
    in
        Extracted,

    // Extract Unit directly
    Unit = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{0, RelativePosition.FromEnd}), 
                    "\")
    in
        Extracted,

    // Extract Date directly
    Date = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted = Text.BeforeDelimiter(
                        Text.AfterDelimiter(
                            Text.BeforeDelimiter(
                                AfterDocuments, "\2. Calculations", {0, RelativePosition.FromEnd}),
                        "\",{2, RelativePosition.FromEnd}), // Extract the unit before the last delimiter
                    "\")
    in
        Extracted,


            // Extract Dole directly
    Role = 
    let
        Source = FilePathUrl,
        FilePath = Source[FilePath]{0}, // Extract the first row's FilePath value
        AfterDocuments = Text.AfterDelimiter(FilePath, "Documents\"), // Extract everything after "Documents\"
        Extracted =                       
                            Text.AfterDelimiter(Text.BeforeDelimiter(AfterDocuments, "\[", {0, RelativePosition.FromEnd}),"\",{0, RelativePosition.FromEnd})
                       
    in
        Extracted,

    // Extract Filename
    FileName = 
    let
        Source = FilePathUrl,
        #"Extracted Text After Delimiter" = Table.TransformColumns(Source, {{"FilePath", each Text.AfterDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
        #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Extracted Text After Delimiter", {{"FilePath", each Text.BeforeDelimiter(_, "]"), type text}}),
        #"Replaced Value" = Table.ReplaceValue(#"Extracted Text Before Delimiter","[","",Replacer.ReplaceText,{"FilePath"}),
        String = #"Replaced Value"{0}[FilePath]
    in 
        String,

    // PART D - Table
    // Create the table with variable names and their corresponding values
    TABLE = #table(
        {"Variable Name", "Value"},
        {
            {"UserName", UserName},
            {"Root Path", IMPORTRootPath{0}[FilePath]},
            {"FilePathUrl", FilePathUrl{0}[FilePath]},
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

shared FilePath = let
    Source = RolePath,
    #"Converted to Table" = #table(1, {{Source}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Converted to Table", {{"Column1", each Text.BeforeDelimiter(_, "\", {1, RelativePosition.FromEnd}), type text}}),
    Column1 = #"Extracted Text Before Delimiter"{0}[Column1]
in
    Column1;

shared Role = let
    Source = RolePathTABLE,
    Value = Source{6}[Value]
in
    Value;