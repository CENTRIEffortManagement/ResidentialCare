// Power Query from: CapacityDistrib(A.1)-shifts - enhance1.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\RN\CapacityDistrib(A.1)-shifts - enhance1.xlsx
// Extracted: 2026-05-21T00:53:48.926Z

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
shared #"AppliedAllocation+ActiveDays" = let
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
    // 1) Start from your per‑shift allocations
    Raw = #"AppliedAllocation+ActiveDays",

    // 2) Keep only Resource, Day and the raw Allocation
    Base = Table.SelectColumns(Raw,{"Resource", "Day", "Allocation"}),

    // 3) Turn each shift into a 1/0 flag
    ShiftFlagged = Table.TransformColumns(
        Base,
        {{"Allocation", each if _ <> null and _ > ValidAllocation then 1 else 0, Int64.Type}}
    ),

    // 4) Collapse to one row per Resource/Day (1 if any shift had allocation)
    Daily = Table.Group(ShiftFlagged, {"Resource", "Day"}, {{"AllocFlag", each List.Max([Allocation]), type number}}),

    // 5) Buffer that small daily table once
    LeanDailyBUFFER = Table.Buffer(Daily),

    // 6) Sort & group so we can index within each Resource
    Sorted = Table.Sort(LeanDailyBUFFER, {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    Grouped = Table.Group(
        Sorted,
        {"Resource"},
        {{"AllData", each Table.AddIndexColumn(_, "Idx", 0, 1, Int64.Type), type table}}
    ),

    // 7) **Note the double‐braces** here—this makes your transformOperations a list of one operation
    WithCodes = Table.TransformColumns(
        Grouped,
        {
            {
                "AllData",
                (tbl) =>
                    let
                        flags    = Table.Column(tbl, "AllocFlag"),
                        n        = List.Count(flags),
                        getD     = (i, o) => if i+o < 0 or i+o >= n then "0" else Text.From(flags{i+o}),
                        enriched = Table.AddColumn(
                                     tbl,
                                     "Code",
                                     each Text.Combine(List.Transform({-2..2}, (off) => getD([Idx], off)), "")
                                   )
                    in
                        enriched
            }
        }
    ),

    // 8) Expand only the new fields (Day, AllocFlag, Idx, Code)
    Expanded = Table.ExpandTableColumn(
        WithCodes,
        "AllData",
        {"Day", "AllocFlag", "Idx", "Code"},
        {"Day", "AllocFlag", "Idx", "Code"}
    ),
    #"Renamed Columns" = Table.RenameColumns(Expanded,{{"Idx", "ResIndex"}})
in
    #"Renamed Columns"
;

[ Description = "BUFFER" ]
shared AllocationCode = let
    // 1) Base daily codes
    SourceDaily = ResDaysAllocated,  // must exist already

    // 2) Group by Resource → sort by Day → add a zero‑based index
    Grouped = Table.Group(
      SourceDaily,
      {"Resource"},
      {{"Group",
        each Table.AddIndexColumn(
               Table.Sort(_, {"Day", Order.Ascending}),
               "Idx", 0, 1, Int64.Type
             ),
        type table
      }}
    ),

    // 3) Inside each Resource‑group, peek at Resource IDs ±2 days
    WithNeighbors = Table.TransformColumns(
      Grouped,
      {"Group",(tbl) =>
        let
          resList = Table.Column(tbl, "Resource"),
          n       = List.Count(resList),
          getRes  = (i,off) => let p = i+off in if p<0 or p>=n then null else resList{p},
          step1   = Table.AddColumn(tbl, "Yesterday.Resource2", each getRes([Idx], -2), Int64.Type),
          step2   = Table.AddColumn(step1,   "Yesterday.Resource1", each getRes([Idx], -1), Int64.Type),
          step3   = Table.AddColumn(step2,   "Tomorrow.Resource1",  each getRes([Idx],  1), Int64.Type),
          step4   = Table.AddColumn(step3,   "Tomorrow.Resource2",  each getRes([Idx],  2), Int64.Type)
        in
          step4
      }
    ),

    // 4) Flatten back to one table (bringing in the new neighbor columns)
    Expanded = Table.ExpandTableColumn(WithNeighbors, "Group", {"Day", "AllocFlag", "ResIndex", "Code", "Yesterday.Resource2", "Yesterday.Resource1", "Tomorrow.Resource1", "Tomorrow.Resource2"}, {"Day", "AllocFlag", "ResIndex", "Code", "Yesterday.Resource2", "Yesterday.Resource1", "Tomorrow.Resource1", "Tomorrow.Resource2"}),

    // 5) Original “Added RESOURCEEND” logic, unchanged
    AddedResourceEnd = Table.AddColumn(
      Expanded,
      "ResourceStartEnd",
      each 
        if      [Resource] <> [Tomorrow.Resource1]   then "ResourceEnd1"
        else if [Resource] <> [Yesterday.Resource1]  then "ResourceStart1"
        else if [Resource] <> [Tomorrow.Resource2]   then "ResourceEnd2"
        else if [Resource] <> [Yesterday.Resource2]  then "ResourceStart2"
        else null,
      type text
    ),





 //#"Removed Other Columns1" = Table.SelectColumns(#"Added RESOURCEEND",{"Resource", "Day", "Allocation", "ResIndex", "ResDayAllocated", "Tomorrow.Resource1", "Code", "ResourceStartEnd"}),
    #"Added CLUSTERPOINTS" = Table.AddColumn(AddedResourceEnd, "ClusterPoints", each if (                    
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
                and [AllocFlag] <> 0
                )
           
        then "Allocated-ClusterEnd" 

        else 
            if 
            //General end
            [Code] = "00100" 
            then "Allocated-SingleStart/End"        
        else null),
    BUFFER = Table.Buffer(#"Added CLUSTERPOINTS")
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
    #"Renamed Columns" = Table.RenameColumns(#"Filled Down",{{"ClusterIndex", "ClusterIndexX"}, {"AllocFlag", "Allocation"}}),
    #"Added NO ALLOCATION NULL" = Table.AddColumn(#"Renamed Columns", "ClusterIndex", each if [Allocation] = 0 then null else [ClusterIndexX]),
    #"Filled Up" = Table.FillUp(#"Added NO ALLOCATION NULL",{"ClusterIndex"}),
    #"Removed Columns" = Table.RemoveColumns(#"Filled Up",{"ClusterIndexX", "ClusterPoints"}),

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
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Allocation", Int64.Type}}),
    #"Grouped CLUSTERSIZE" = Table.Group(#"Changed Type", {"ClusterIndex"}, {{"ClusterAllocation", each List.Sum([Allocation]), type nullable number}}),
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
    #"Added NWD TYPE" = Table.AddColumn(#"Expanded ClusterAllocation", "NWDType", each if [Allocation] = 1  then "Allocated" 

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

shared RolePathTABLE = // Version 25.02 flexible ResidentialCare
let
    FilePathUrl =
    let
        Source = try Excel.CurrentWorkbook(){[Name="FilePAthUrl"]}[Content] otherwise Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        FirstColumn = Table.ColumnNames(Source){0},
        RenamedColumns = if FirstColumn = "FilePath" then Source else Table.RenameColumns(Source, {{FirstColumn, "FilePath"}}, MissingField.Ignore),
        ReplacedValue = Table.TransformColumns(RenamedColumns, {{"FilePath", each Text.Replace(Text.From(_), "/", "\"), type text}}),
        BufferedTable = Table.Buffer(ReplacedValue)
    in
        BufferedTable,

    RawFilePath = FilePathUrl{0}[FilePath],
    CentriSyncPaths_Source = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    CentriSyncPaths_Table = CentriSyncPaths_Source{[Item="CentriSyncPaths",Kind="Table"]}[Data],
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    FilePath = NormalizePath(RawFilePath),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
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
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    BestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0} else error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    RootPath = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Segments = List.Select(Text.Split(RootPath, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare"),
    UnitsIndex = List.PositionOf(Segments, "UNITS"),
    CalcIndex = List.PositionOf(Segments, "2. Calculations"),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    Unit = if UnitsIndex >= 0 and List.Count(Segments) > UnitsIndex + 1 then Segments{UnitsIndex + 1} else null,
    Role = if CalcIndex >= 0 and List.Count(Segments) > CalcIndex + 1 then Segments{CalcIndex + 1} else null,
    FileName = try Text.BetweenDelimiters(LocalFullPath, "[", "]") otherwise Text.AfterDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
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

shared FilePath = let
    Source = RolePath,
    #"Converted to Table" = #table(1, {{Source}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Converted to Table", {{"Column1", each Text.BeforeDelimiter(_, "\", {1, RelativePosition.FromEnd}), type text}}),
    Column1 = #"Extracted Text Before Delimiter"{0}[Column1]
in
    Column1;

shared Role = let
    Source = RolePathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Role")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;
