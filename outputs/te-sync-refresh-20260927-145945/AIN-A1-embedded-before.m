section Section1;

// Query: IMPORT Demand
// Purpose: Open Demand.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT Demand" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\Demand.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: Demand Source Prepare
// Purpose: Select and type the role-specific shift-demand table used by A.1.
shared #"Demand Source Prepare" = let
    Navigation = #"IMPORT Demand",
    Matches = Table.SelectRows(Navigation, each [Item] = "ShiftDemandHCAverageANACC" and [Kind] = "Table"),
    DemandTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 demand import", "Expected exactly one ShiftDemandHCAverageANACC table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Facility", "Role", "Date", "ShiftPeriod", "UnitShiftEffort", "ShiftDurations.Duration", "ShiftDemandHCAverage"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(DemandTable)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then DemandTable
        else error Error.Record("A.1 demand import", "ShiftDemandHCAverageANACC is missing required columns.", [MissingColumns = MissingColumns]),
    FilteredRole = Table.SelectRows(ValidatedTable, each [Role] = Role),
    RenamedShiftDuration = Table.RenameColumns(FilteredRole, {{"ShiftDurations.Duration", "ShiftDurations.ShiftDuration"}}),
    Typed = Table.TransformColumnTypes(RenamedShiftDuration, {{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"ShiftPeriod", type text}, {"UnitShiftEffort", type number}, {"ShiftDurations.ShiftDuration", type number}, {"ShiftDemandHCAverage", type number}})
in
    Typed;

// Query: IMPORT ShiftUnitDemandHRS !!
// Purpose: Preserve the existing prepared-demand interface for downstream queries.
shared #"IMPORT ShiftUnitDemandHRS !!" = #"Demand Source Prepare";

// Query: IMPORT StaffListMaster
// Purpose: Open StaffListMaster.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT StaffListMaster" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\StaffListMaster.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: Masterlist Prepare
// Purpose: Validate and prepare the role-specific employee contract used by A.1.
// Output: One row per source master-list record before Resource-grain validation.
shared #"Masterlist Prepare" = let
    Navigation = #"IMPORT StaffListMaster",
    Matches = Table.SelectRows(Navigation, each [Item] = "Table_Masterlist" and [Kind] = "Table"),
    MasterlistTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 master-list import", "Expected exactly one Table_Masterlist table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Name", "Role", "Resource", "EmployeeID", "Facility-Abbrev", "PreferredRole", "Contracted FN Hours",
        "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(MasterlistTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(MasterlistTable, RequiredColumns)
        else error Error.Record("A.1 master-list import", "Table_Masterlist is missing required contract columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected,
        {{"Name", type text}, {"Role", type text}, {"Resource", Int64.Type}, {"EmployeeID", type text},
         {"Facility-Abbrev", type text}, {"PreferredRole", type text}, {"Contracted FN Hours", type nullable number},
         {"Contracted Shifts", type nullable number}, {"Settings Max Availability", type nullable number},
         {"Effective Shift Cap", type nullable number}, {"Limit Basis", type text}}),
    FilteredRole = Table.SelectRows(Typed, each [Role] = Role)
in
    FilteredRole;

// Query: IMPORT Masterlist !!
// Purpose: Preserve the existing prepared-master-list interface for downstream queries.
shared #"IMPORT Masterlist !!" = #"Masterlist Prepare";

// Query: A1CheckResult
// Purpose: Convert a validation count or evaluation error into a consistent check row.
shared A1CheckResult = (name as text, evaluate as function) as record =>
    let
        Result = try evaluate()
    in
        [Check = name, Status = if Result[HasError] then "Fail" else if Result[Value] = 0 then "Pass" else "Fail",
         Failures = if Result[HasError] then null else Result[Value],
         Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null];

// Query: IMPORT Capacity ShiftAvailability
// Purpose: Open Capacity-ShiftAvailability.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT Capacity ShiftAvailability" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\Capacity-ShiftAvailability.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: ResDayShift Prepare
// Purpose: Validate and type only the availability fields consumed by A.1, then apply the dynamic role scope.
shared #"ResDayShift Prepare" = let
    Navigation = #"IMPORT Capacity ShiftAvailability",
    Matches = Table.SelectRows(Navigation, each [Item] = "ResDayShift" and [Kind] = "Table"),
    ResDayShiftTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 availability import", "Expected exactly one ResDayShift table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Name", "Role", "Week", "Day", "Shift", "EffectiveShiftHrs", "Date"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(ResDayShiftTable)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then ResDayShiftTable
        else error Error.Record("A.1 availability import", "ResDayShift is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(ValidatedTable, {{"Name", type text}, {"Role", type text}, {"Week", Int64.Type}, {"Day", type text}, {"Shift", type text}, {"EffectiveShiftHrs", type number}, {"Date", type date}}),
    FilteredRole = Table.SelectRows(Typed, each [Role] = Role)
in
    FilteredRole;

// Query: IMPORT ResDayShift !!
// Purpose: Preserve the existing prepared-availability interface for downstream queries.
shared #"IMPORT ResDayShift !!" = #"ResDayShift Prepare";

// Query: IMPORT AllocationByShiftAverage
// Purpose: Open AllocationByShiftAverage.xlsx once and expose its workbook navigation table without business transformations.
shared #"IMPORT AllocationByShiftAverage" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\AllocationByShiftAverage.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: ResourceShiftAllocation Prepare
// Purpose: Validate and prepare the role-specific resource-shift allocations used by A.1.
shared #"ResourceShiftAllocation Prepare" = let
    Navigation = #"IMPORT AllocationByShiftAverage",
    Matches = Table.SelectRows(Navigation, each [Item] = "ResourceShiftAllocation" and [Kind] = "Table"),
    AllocationTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("A.1 allocation import", "Expected exactly one ResourceShiftAllocation table.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"ShiftDate", "ShiftPeriod", "IntervalAssociatedShift", "Name", "Role", "ResShiftEffort", "ResShiftEffectiveRatio", "ResShiftFTE"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(AllocationTable)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then AllocationTable
        else error Error.Record("A.1 allocation import", "ResourceShiftAllocation is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(ValidatedTable, {{"ShiftDate", type date}, {"ShiftPeriod", type text}, {"IntervalAssociatedShift", type text}, {"Name", type text}, {"Role", type text}, {"ResShiftEffort", type number}, {"ResShiftEffectiveRatio", type number}, {"ResShiftFTE", type number}}),
    FilteredRole = Table.SelectRows(Typed, each [Role] = Role)
in
    FilteredRole;

// Query: IMPORT ResourceShiftAllocation - Role!!
// Purpose: Preserve the existing prepared-allocation interface for downstream queries.
shared #"IMPORT ResourceShiftAllocation - Role!!" = #"ResourceShiftAllocation Prepare";

shared MaxShiftCluster = 5 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

[ Description = "Use to filter out shift creep into adjacent shifts" ]
shared NotShiftThreshold = 0.2 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared #"SET NWDPriority" = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("JYopEsAgEAS/Qq1ew+b28aGCpBA5QFFx/D/L4Lp7JgSyxHSmtz7XXRJFDiRa7G6OnKGDqhc2TqCjqhP2FjZhBM4YgEur/bC22nFT/GopFOMP", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [NWDPriority = _t, #"NWD Type" = _t])
in
    Source;

[ Description = "BUFFER" ]
// Query: PeriodShiftDay B
// Purpose: Buffer the role-scoped Settings permutation dimensions used throughout A.1.
shared #"PeriodShiftDay B" = let
    Source = #"PermutationDimensions Prepare",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Date", "Day", "Shifts", "Period", "RolesList"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

// Query: Resources
// Purpose: Retain only the stable fields required to map role data to Resource grain.
shared Resources = let
    Source = #"IMPORT Masterlist !!",
    Selected = Table.SelectColumns(Source, {"Name", "Role", "Resource"}),
    Buffered = Table.Buffer(Selected)
in
    Buffered;

// Query: A1ResourceIdentity_CHECK
// Purpose: Prevent name-based availability or allocation joins from multiplying or crossing Resources.
shared A1ResourceIdentity_CHECK = let
    ResourceMap = Resources,
    AvailabilityNames = Table.Distinct(Table.SelectColumns(#"IMPORT ResDayShift !!", {"Name", "Role"})),
    AllocationNames = Table.Distinct(Table.SelectColumns(#"IMPORT ResourceShiftAllocation - Role!!", {"Name", "Role"})),
    AvailabilityCoverage = Table.NestedJoin(AvailabilityNames, {"Name", "Role"}, ResourceMap, {"Name", "Role"}, "ResourceMatches", JoinKind.LeftOuter),
    AvailabilityMatchCounts = Table.AddColumn(AvailabilityCoverage, "MatchCount", each Table.RowCount([ResourceMatches]), Int64.Type),
    AllocationCoverage = Table.NestedJoin(AllocationNames, {"Name", "Role"}, ResourceMap, {"Name", "Role"}, "ResourceMatches", JoinKind.LeftOuter),
    AllocationMatchCounts = Table.AddColumn(AllocationCoverage, "MatchCount", each Table.RowCount([ResourceMatches]), Int64.Type),
    Checks = {
        A1CheckResult("Resource identifiers are populated", () => Table.RowCount(Table.SelectRows(ResourceMap, each [Resource] = null))),
        A1CheckResult("Resource identifiers are unique", () => Table.RowCount(ResourceMap) - Table.RowCount(Table.Distinct(ResourceMap, {"Resource"}))),
        A1CheckResult("Resource names are populated", () => Table.RowCount(Table.SelectRows(ResourceMap, each [Name] = null or Text.Trim([Name]) = ""))),
        A1CheckResult("Resource names are unique within role", () => Table.RowCount(ResourceMap) - Table.RowCount(Table.Distinct(ResourceMap, {"Name", "Role"}))),
        A1CheckResult("Availability names map to exactly one Resource", () => Table.RowCount(Table.SelectRows(AvailabilityMatchCounts, each [MatchCount] <> 1))),
        A1CheckResult("Allocation names map to exactly one Resource", () => Table.RowCount(Table.SelectRows(AllocationMatchCounts, each [MatchCount] <> 1)))
    },
    CheckTable = Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    Buffered = Table.Buffer(CheckTable)
in
    Buffered;

// Query: A1ResourceContract
// Purpose: Expose the role-filtered master-list contract at one row per Resource for the A.1 cap join.
shared A1ResourceContract = let
    Source = #"IMPORT Masterlist !!",
    Selected = Table.SelectColumns(Source,
        {"Resource", "Role", "EmployeeID", "Facility-Abbrev", "PreferredRole", "Contracted FN Hours",
         "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"})
in Table.Buffer(Selected);

// Query: ResourcePeriodTABLE-empty
// Purpose: Build the Resource-period skeleton while standardising Settings input Shifts to downstream Shift.
shared #"ResourcePeriodTABLE-empty" = let
    Source = Resources,
    #"Added Custom" = Table.AddColumn(Source, "Periods", each #"PeriodShiftDay B"),
    // Settings exposes Shifts; downstream A.1 calculations use the singular Shift interface.
    #"Expanded Periods" = Table.ExpandTableColumn(#"Added Custom", "Periods", {"Period", "Shifts", "Day"}, {"Period", "Shift", "Day"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Periods", each ([Period] <> "W")),
    #"Changed Type" = Table.TransformColumnTypes(#"Filtered Rows",{{"Resource", Int64.Type}, {"Period", Int64.Type}})
in
    #"Changed Type";

[ Description = "BUFFER" ]
// Query: ResPeriodAvailabilityTABLE !!
// Purpose: Map role availability to Resource periods after the identity checks pass.
shared #"ResPeriodAvailabilityTABLE !!" = let
    IdentityFailures = Table.SelectRows(A1ResourceIdentity_CHECK, each [Status] <> "Pass"),
    Source = if Table.IsEmpty(IdentityFailures) then #"IMPORT ResDayShift !!"
        else error Error.Record("A.1 resource identity validation", "Availability mapping requires unique Resource identities.", IdentityFailures),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Name", "Role", "Week", "Shift", "Date"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Removed Other Columns",{{"Date", type date}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type2", {"Name"}, Resources, {"Name"}, "Resources", JoinKind.LeftOuter),
    #"Expanded Resources" = Table.ExpandTableColumn(#"Merged Queries", "Resources", {"Resource"}, {"Resource"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded Resources", {"Date", "Shift", "Role"}, #"PeriodShiftDay B", {"Date", "Shifts", "RolesList"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodShiftDay", {"Period"}, {"Period"}),
    // Make Availability dynamic
    #"Added AVAILABILITY TEMP" = Table.AddColumn(#"Expanded PeriodShiftDay", "Availability", each 1),
    #"Removed Other Columns1" = Table.SelectColumns(#"Added AVAILABILITY TEMP",{"Role", "Resource", "Period", "Availability"}),
    // Buffer
    BUFFER = Table.Buffer(#"Removed Other Columns1")
in
    BUFFER;

shared ResDayPeriodAvailabilityTABLE = let
    Source = #"ResPeriodAvailabilityTABLE !!",
    #"Merged Queries" = Table.NestedJoin(Source, {"Period", "Role"}, #"PeriodShiftDay B", {"Period", "RolesList"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries", "PeriodShiftDay", {"Day"}, {"Day"}),
    BUFFER = Table.Buffer(#"Expanded PeriodShiftDay")
in
    BUFFER;

shared ResDayAvailabilityTABLE = let
    Source = ResDayPeriodAvailabilityTABLE,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Resource", "Day"}, {{"DayAvailbility", each List.Sum([Availability]), type number}, {"DayShiftCount", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

shared ResPeriodAvailabilitySUM = let
    Source = List.Sum(#"ResPeriodAvailabilityTABLE !!"[Availability])
in
    Source;

// Query: ResPeriodAllocationCHECK
// Purpose: Map role allocations to Resources after the identity checks pass.
shared ResPeriodAllocationCHECK = let
    IdentityFailures = Table.SelectRows(A1ResourceIdentity_CHECK, each [Status] <> "Pass"),
    Source = if Table.IsEmpty(IdentityFailures) then #"IMPORT ResourceShiftAllocation - Role!!"
        else error Error.Record("A.1 resource identity validation", "Allocation mapping requires unique Resource identities.", IdentityFailures),
    #"Filtered DATEFROM" = Table.SelectRows(Source, each ([ShiftDate] >= #"EXTRACT Date_From")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered DATEFROM",{"ResShiftEffort", "ResShiftEffectiveRatio"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Name", "Role"}, Resources, {"Name", "Role"}, "Resources", JoinKind.LeftOuter),
    #"Expanded Resources" = Table.ExpandTableColumn(#"Merged Queries", "Resources", {"Resource"}, {"Resource"}),
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Resources",{{"ResShiftFTE", "Allocation"}})
in
    #"Renamed Columns";

[ Description = "BUFFER" ]
shared #"ResPeriodAllocation-B" = let
    Source = ResPeriodAllocationCHECK,
    #"Filtered Rows NULL" = Table.SelectRows(Source, each ([Resource] <> null)),
    #"Added SHIFTNUMBER" = Table.AddColumn(#"Filtered Rows NULL", "ShiftNumber", each if [IntervalAssociatedShift] = "AM" then 1 else if [IntervalAssociatedShift] = "PM" then 2 else if [IntervalAssociatedShift] = "NIGHT" then 3 else null),
    BUFFER = Table.Buffer(#"Added SHIFTNUMBER")
in
    BUFFER;

shared #"ResPeriodAllocation-ShiftBias" = let
    Source = #"ResPeriodAllocation-B",
    #"Grouped Rows1" = Table.Group(Source, {"IntervalAssociatedShift", "Resource", "ShiftDate"}, {{"Allocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped Rows1", each ([Allocation] > NotShiftThreshold)),
    #"Added SHIFTINDEX" = Table.AddColumn(#"Filtered Rows", "ShiftIndex", each if [IntervalAssociatedShift] = "AM" then 1 else if [IntervalAssociatedShift] = "PM" then 2 else 3),
    #"Grouped SHIFTBIAS" = Table.Group(#"Added SHIFTINDEX", {"Resource"}, {{"ShiftBias", each List.Average([ShiftIndex]), type number}}),
    BUFFER = Table.Buffer(#"Grouped SHIFTBIAS")
in
    BUFFER;

[ Description = "BUFFER" ]
shared ResPeriodAllocationTABLE = let
    Source = #"ResPeriodAllocation-B",
    #"Merged Queries1" = Table.NestedJoin(Source, {"ShiftDate", "IntervalAssociatedShift"}, #"PeriodShiftDay B", {"Date", "Shifts"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodShiftDay", {"Period"}, {"Period"}),
    #"Grouped ALLOCATION" = Table.Group(#"Expanded PeriodShiftDay", {"ShiftDate", "IntervalAssociatedShift", "Role", "Resource", "Period", "ShiftNumber"}, {{"Allocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered NOTFULLSHIFT" = Table.SelectRows(#"Grouped ALLOCATION", each ([Allocation] > NotShiftThreshold)),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered NOTFULLSHIFT",{"Role", "Resource", "Period", "ShiftNumber", "Allocation"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

shared ResDayAllocationTABLE = let
    Source = Table.NestedJoin(ResPeriodAllocationTABLE, {"Period"}, #"PeriodShiftDay B", {"Period"}, "PeriodShiftDay", JoinKind.LeftOuter),
    #"Expanded PeriodShiftDay1" = Table.ExpandTableColumn(Source, "PeriodShiftDay", {"Day"}, {"Day"}),
    #"Group RESDAYALLOCATION" = Table.Group(#"Expanded PeriodShiftDay1", {"Resource", "Day"}, {{"ResDayAllocation", each List.Sum([Allocation]), type nullable number}, {"ShiftAverage", each List.Average([ShiftNumber]), type number}}),
    BUFFER = Table.Buffer(#"Group RESDAYALLOCATION")
in
    BUFFER;

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
    #"Grouped C'" = Table.Group(#"Filtered Rows", {"Period"}, {{"PeriodC'", each List.Sum([Availability]), type number}})
in
    #"Grouped C'";

// Query: PeriodC'-D
// Purpose: Compare potential capacity with demand and expose a nullable diagnostic capacity-to-demand ratio.
shared #"PeriodC'-D" = let
    Source = #"PeriodC'",
    Custom1 = Table.Buffer(Source),
    #"Merged Queries" = Table.NestedJoin(Custom1, {"Period"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries", "PeriodDemandTABLE", {"D"}, {"D"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C'-D", each [#"PeriodC'"] - [D], type number),
    // The ratio is diagnostic only; null preserves undefined or invalid demand without propagating an error to A.2.
    #"Inserted Safe Division" = Table.AddColumn(#"Inserted Subtraction", "C'/D", each
        let
            Demand = try Number.From([D]) otherwise null
        in
            if Demand = null or Demand = 0 then null else [#"PeriodC'"] / Demand,
        type nullable number
    )
in
    #"Inserted Safe Division";

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

// Query: A1ResourceContract_CHECK
// Purpose: Require one valid preferred-role contract for every Resource contributing role availability.
shared A1ResourceContract_CHECK = let
    Contracts = A1ResourceContract,
    AvailabilityResources = Table.Distinct(Table.SelectColumns(#"ResC'", {"Resource"})),
    Coverage = Table.NestedJoin(AvailabilityResources, {"Resource"}, Contracts, {"Resource"}, "Contract", JoinKind.LeftOuter),
    WithMatchCount = Table.AddColumn(Coverage, "ContractMatchCount", each Table.RowCount([Contract]), Int64.Type),
    SingleMatches = Table.SelectRows(WithMatchCount, each [ContractMatchCount] = 1),
    ExpandedMatches = Table.ExpandTableColumn(SingleMatches, "Contract",
        {"Role", "EmployeeID", "PreferredRole", "Effective Shift Cap", "Limit Basis"},
        {"Contract Role", "EmployeeID", "PreferredRole", "Effective Shift Cap", "Limit Basis"}),
    Checks = {
        A1CheckResult("Contract Resources are populated", () => Table.RowCount(Table.SelectRows(Contracts, each [Resource] = null))),
        A1CheckResult("Unique Resource contracts", () => Table.RowCount(Contracts) -
            Table.RowCount(Table.Distinct(Contracts, {"Resource"}))),
        A1CheckResult("Complete availability contract coverage", () =>
            Table.RowCount(Table.SelectRows(WithMatchCount, each [ContractMatchCount] <> 1))),
        A1CheckResult("Valid effective shift caps", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [Effective Shift Cap] = null or [Effective Shift Cap] <= 0
            or [Effective Shift Cap] <> Number.RoundDown([Effective Shift Cap])))),
        A1CheckResult("Preferred role matches contract", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [Contract Role] <> Role or [PreferredRole] <> Role))),
        A1CheckResult("Employee identifiers are populated", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [EmployeeID] = null or (try Text.Trim([EmployeeID]) otherwise "") = ""))),
        A1CheckResult("Limit basis is populated", () => Table.RowCount(Table.SelectRows(ExpandedMatches, each
            [Limit Basis] = null or (try Text.Trim([Limit Basis]) otherwise "") = "")))
    },
    CheckTable = Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    Buffered = Table.Buffer(CheckTable)
in
    Buffered;

// Query: ResAv-C'
// Purpose: Apply the employee contract cap to redistributable roster availability at Resource grain.
shared #"ResAv-C'" = let
    IdentityFailures = Table.SelectRows(A1ResourceIdentity_CHECK, each [Status] <> "Pass"),
    ContractFailures = if Table.IsEmpty(IdentityFailures) then Table.SelectRows(A1ResourceContract_CHECK, each [Status] <> "Pass")
        else #table(type table [Check = text, Status = text, Failures = nullable number, Details = nullable text], {}),
    Source = if not Table.IsEmpty(IdentityFailures) then
            error Error.Record("A.1 resource identity validation", "Required Resource identity checks failed.", IdentityFailures)
        else if not Table.IsEmpty(ContractFailures) then
            error Error.Record("A.1 resource contract validation", "Required Resource contract checks failed.", ContractFailures)
        else #"ResC'",
    // Join once at Resource grain; never repeat roster contract totals across Resource-period rows.
    JoinedContract = Table.NestedJoin(Source, {"Resource"}, A1ResourceContract, {"Resource"}, "ResourceContract", JoinKind.LeftOuter),
    ExpandedContract = Table.ExpandTableColumn(JoinedContract, "ResourceContract",
        {"Effective Shift Cap", "Limit Basis"}, {"Effective Shift Cap", "Limit Basis"}),
    #"Added AVAILCAP" = Table.AddColumn(ExpandedContract, "RosterAvailabilityCAPPED", each
        if [RosterAvailability] > [Effective Shift Cap] then [Effective Shift Cap] else [RosterAvailability], type number),
    #"Inserted REDUCTION" = Table.AddColumn(#"Added AVAILCAP", "AvailabilityReduction", each [RosterAvailability] - [RosterAvailabilityCAPPED], type number),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted REDUCTION",{{"AvailabilityReduction", "ResAv-C'"}})
in
    #"Renamed Columns";

[ Description = "BUFFER   Tag RP cell with R that need to reduced becuase over allocated" ]
shared #"ResPeriod(ExcessPot)SpareAvailabilityTABLE" = let
    Source = #"C' ResSingleShiftDayTABLE",
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Period"}, ResPeriodAllocationTABLE, {"Resource", "Period"}, "ResPeriodAllocationTABLE", JoinKind.LeftOuter),
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
    #"Filtered Rows" = Table.SelectRows(Source, each ([UnassignedAvail] > 0)),
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

// Query: ResDaysAllocated
// Purpose: Collapse shift-level allocations to daily flags and calculate deterministic five-day allocation codes per Resource.
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

    // 5) Enforce numeric chronological order before buffering and sequence calculations.
    TypedDay = Table.TransformColumnTypes(Daily, {{"Day", Int64.Type}}),
    LeanDailyBUFFER = Table.Buffer(TypedDay),

    // 6) Sort each nested Resource table explicitly; Table.Group does not guarantee preservation of its input row order.
    Sorted = Table.Sort(LeanDailyBUFFER, {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    Grouped = Table.Group(
        Sorted,
        {"Resource"},
        {{"AllData", each Table.AddIndexColumn(
            Table.Sort(_, {{"Day", Order.Ascending}}),
            "Idx", 0, 1, Int64.Type
        ), type table}}
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
// Query: AllocationCode
// Purpose: Identify cluster boundaries from deterministic Resource/day allocation sequences.
shared AllocationCode = let
    // 1) Base daily codes
    Source = ResDaysAllocated,
    // Cluster neighbour offsets must use numeric day order, never lexical text order.
    SourceDaily = Table.TransformColumnTypes(Source, {{"Day", Int64.Type}}),

    // 2) Group by Resource → sort by Day → add a zero‑based index
    Grouped = Table.Group(
      SourceDaily,
      {"Resource"},
      {{"Group",
        each Table.AddIndexColumn(
               Table.Sort(_, {{"Day", Order.Ascending}}),
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
    SortedExpanded = Table.Sort(Expanded, {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),

    // 5) Original “Added RESOURCEEND” logic, unchanged
    AddedResourceEnd = Table.AddColumn(
      SortedExpanded,
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
    // A one-day allocation is both a start and an end, including when it occurs on a Resource boundary.
    #"Added CLUSTERPOINTS" = Table.AddColumn(AddedResourceEnd, "ClusterPoints", each if [Code] = "00100" then "Allocated-SingleStart/End"
        else if (
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

        else null),
    BUFFER = Table.Buffer(#"Added CLUSTERPOINTS")
in
    BUFFER;

// Query: ClusterIDStart
// Purpose: Assign stable sequence numbers to every cluster start, including single-day clusters.
shared ClusterIDStart = let
    Source = Table.Sort(#"AllocationCode", {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Day", "ClusterPoints"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each [ClusterPoints] = "Allocated-ClusterStart" or [ClusterPoints] = "Allocated-ClusterSINGLE" or [ClusterPoints] = "Allocated-SingleStart/End"),
    #"Added Index1" = Table.AddIndexColumn(#"Filtered Rows", "ClusterIndex", 1, 1, Int64.Type)
in
    #"Added Index1";

// Query: ClusterIDEnd
// Purpose: Assign matching sequence numbers to every cluster end, including single-day clusters.
shared ClusterIDEnd = let
    Source = Table.Sort(#"AllocationCode", {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Day", "ClusterPoints"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each [ClusterPoints] = "Allocated-ClusterEnd" or [ClusterPoints] = "Allocated-ClusterSINGLE" or [ClusterPoints] = "Allocated-SingleStart/End"),
    #"Added Index1" = Table.AddIndexColumn(#"Filtered Rows", "ClusterIndex", 1, 1, Int64.Type)
in
    #"Added Index1";

shared #"ClusterIDStart+End" = let
    Source = Table.Combine({ClusterIDEnd, ClusterIDStart}),
    #"Removed Duplicates" = Table.Distinct(Source)
in
    #"Removed Duplicates";

// Query: A1ClusterLookup_CHECK
// Purpose: Validate the Resource-scoped lookup grain and the paired cluster-boundary sequence.
shared A1ClusterLookup_CHECK = let
    LookupKeys = Table.SelectColumns(#"AllocationCode", {"Resource", "ResIndex"}),
    StartCounts = Table.Group(ClusterIDStart, {"Resource"}, {{"StartCount", each Table.RowCount(_), Int64.Type}}),
    EndCounts = Table.Group(ClusterIDEnd, {"Resource"}, {{"EndCount", each Table.RowCount(_), Int64.Type}}),
    BoundaryCounts = Table.NestedJoin(StartCounts, {"Resource"}, EndCounts, {"Resource"}, "EndCounts", JoinKind.FullOuter),
    ExpandedBoundaryCounts = Table.ExpandTableColumn(BoundaryCounts, "EndCounts", {"EndCount"}, {"EndCount"}),
    NormalizedBoundaryCounts = Table.ReplaceValue(ExpandedBoundaryCounts, null, 0, Replacer.ReplaceValue, {"StartCount", "EndCount"}),
    Checks = {
        A1CheckResult("Cluster lookup Resources are populated", () => Table.RowCount(Table.SelectRows(LookupKeys, each [Resource] = null or [ResIndex] = null))),
        A1CheckResult("Cluster lookup keys are unique", () => Table.RowCount(LookupKeys) - Table.RowCount(Table.Distinct(LookupKeys, {"Resource", "ResIndex"}))),
        A1CheckResult("Cluster boundaries are balanced within each Resource", () => Table.RowCount(Table.SelectRows(NormalizedBoundaryCounts, each [StartCount] <> [EndCount])))
    },
    CheckTable = Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]),
    Buffered = Table.Buffer(CheckTable)
in
    Buffered;

[ Description = "BUFFER" ]
// Query: Clusters
// Purpose: Assign cluster numbers through a deterministic Resource-scoped index lookup.
shared Clusters = let
    ValidationFailures = Table.SelectRows(A1ClusterLookup_CHECK, each [Status] <> "Pass"),
    Source = if Table.IsEmpty(ValidationFailures) then Table.Sort(#"AllocationCode", {{"Resource", Order.Ascending}, {"Day", Order.Ascending}})
        else error Error.Record("A.1 cluster validation", "Cluster lookup checks failed.", ValidationFailures),
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "Day", "ClusterPoints"}, #"ClusterIDStart+End", {"Resource", "Day", "ClusterPoints"}, "ClusterIDStart+End", JoinKind.LeftOuter),
    #"Expanded ClusterIDStart+End" = Table.ExpandTableColumn(#"Merged Queries", "ClusterIDStart+End", {"ClusterIndex"}, {"ClusterIndex"}),
    #"Grouped for Resource Fill Down" = Table.Group(#"Expanded ClusterIDStart+End", {"Resource"}, {{"Rows", each Table.FillDown(_, {"ClusterIndex"})}}),
    #"Filled Down" = if Table.IsEmpty(#"Expanded ClusterIDStart+End") then #"Expanded ClusterIDStart+End" else Table.Combine(#"Grouped for Resource Fill Down"[Rows]),
    #"Renamed Columns" = Table.RenameColumns(#"Filled Down",{{"ClusterIndex", "ClusterIndexX"}, {"AllocFlag", "Allocation"}}),
    #"Added NO ALLOCATION NULL" = Table.AddColumn(#"Renamed Columns", "ClusterIndex", each if [Allocation] = 0 then null else [ClusterIndexX]),
    #"Grouped for Resource Fill Up" = Table.Group(#"Added NO ALLOCATION NULL", {"Resource"}, {{"Rows", each Table.FillUp(_, {"ClusterIndex"})}}),
    #"Filled Up" = if Table.IsEmpty(#"Added NO ALLOCATION NULL") then #"Added NO ALLOCATION NULL" else Table.Combine(#"Grouped for Resource Fill Up"[Rows]),
    #"Sorted Filled Rows" = Table.Sort(#"Filled Up", {{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Removed Columns" = Table.RemoveColumns(#"Sorted Filled Rows",{"ClusterIndexX", "ClusterPoints"}),

    // Preserve the established Code offsets while representing "00000" as no lookup.
    #"Added Lookup Offset" = Table.AddColumn(#"Removed Columns", "LookupOffset", each
        if [Code] = "00000" then null else
        if List.Contains({"11000", "11011", "01010", "01001", "01000", "11001", "11010"}, [Code]) then -1 else
        if List.Contains({"10000", "10001", "10010"}, [Code]) then -2 else
        if List.Contains({"00011", "00010"}, [Code]) then 1 else
        if List.Contains({"00001", "10011"}, [Code]) then 2 else
        0,
        Int64.Type
    ),
    #"Added Lookup ResIndex" = Table.AddColumn(#"Added Lookup Offset", "LookupResIndex", each
        if [LookupOffset] = null then null else [ResIndex] + [LookupOffset],
        Int64.Type
    ),
    #"Selected Cluster Lookup Columns" = Table.Buffer(Table.SelectColumns(#"Removed Columns", {"Resource", "ResIndex", "ClusterIndex"})),
    #"Joined Resource Cluster" = Table.NestedJoin(
        #"Added Lookup ResIndex",
        {"Resource", "LookupResIndex"},
        #"Selected Cluster Lookup Columns",
        {"Resource", "ResIndex"},
        "ClusterMatch",
        JoinKind.LeftOuter
    ),
    #"Added ClusterNumber" = Table.AddColumn(#"Joined Resource Cluster", "ClusterNumber", each
        if [LookupOffset] = null or Table.IsEmpty([ClusterMatch]) then "X"
        else if Table.RowCount([ClusterMatch]) = 1 then [ClusterMatch]{0}[ClusterIndex]
        else error Error.Record("A.1 cluster lookup", "A Resource and ResIndex matched more than one cluster row.", [Resource = [Resource], ResIndex = [LookupResIndex]])
    ),
    #"Removed Lookup Helpers" = Table.RemoveColumns(#"Added ClusterNumber", {"LookupOffset", "LookupResIndex", "ClusterMatch"}),
    BUFFER = Table.Buffer(#"Removed Lookup Helpers")
in
    BUFFER;

shared ClusterAllocation = let
    Source = Clusters,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Allocation", Int64.Type}}),
    #"Grouped CLUSTERSIZE" = Table.Group(#"Changed Type", {"ClusterIndex"}, {{"ClusterAllocation", each List.Sum([Allocation]), type nullable number}}),
    #"Filtered Rows" = Table.SelectRows(#"Grouped CLUSTERSIZE", each ([ClusterIndex] <> null)),
    BUFFER = Table.Buffer(#"Filtered Rows")
in
    BUFFER;

[ Description = "BUFFER #(lf)S2, P2 dev debt #(lf)Opportunity to discriminate with small Cluster Allocation#(lf)i.e. change the cluster it is associated with" ]
shared #"WD Type" = let
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
    #"Removed Other Columns" = Table.SelectColumns(#"Added NWD TYPE",{"Resource", "Day",  "ClusterNumber", "ClusterAllocation", "NWDType"})
in
    #"Removed Other Columns";

shared ResDayPotentialAvailabilityTABLE = let
    Source = #"WD Type",
  


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
    #"Removed Other Columns" = Table.SelectColumns(#"POTENTIAL ALLOCATION",{"Resource", "Day", "NWDType", "PotentialAvailability"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

[ Description = "BUFFER" ]
shared ResPeriodShiftNWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "R")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"PotentialAvailability"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"NWDType"}, #"SET NWDPriority", {"NWD Type"}, "NWDPrioritiesTABLE", JoinKind.LeftOuter),
    #"Expanded NWDPrioritiesTABLE" = Table.ExpandTableColumn(#"Merged Queries", "NWDPrioritiesTABLE", {"NWDPriority"}, {"NWDPriority"}),
    #"Sorted Rows" = Table.Sort(#"Expanded NWDPrioritiesTABLE",{{"Resource", Order.Ascending}, {"Day", Order.Ascending}}),
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows", {"Day"}, #"PeriodShiftDay B", {"Day"}, "Period.1", JoinKind.LeftOuter),
    #"Expanded PERIODSHIFTS" = Table.ExpandTableColumn(#"Merged Queries1", "Period.1", {"Period", "Shifts"}, {"Period", "Shifts"})
in
    #"Expanded PERIODSHIFTS";

shared ResDayNWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "R")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Resource", "Day", "PotentialAvailability"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

[ Description = "BUFFER" ]
shared ResPeriodWDTABLE = let
    Source = ResDayPotentialAvailabilityTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([PotentialAvailability] = "U")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Resource", "Day", "PotentialAvailability"}),
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

    #"Removed Columns" = Table.RemoveColumns(#"Added ROSTEREDPERIODSTATUS",{"Allocation"}),
    BUFFER = Table.Buffer(#"Removed Columns")
in
    BUFFER;

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
    #"Merged Queries" = Table.NestedJoin(#"Removed MULTIPERIODALLOCATION", {"Resource", "Day"}, ResDayAllocationTABLE, {"Resource", "Day"}, "ResDayAllocationTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayAllocationTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayAllocationTABLE", {"ResDayAllocation"}, {"ResDayAllocation"}),
    #"Filtered UNMACTHED PERIODS REMOVED" = Table.SelectRows(#"Expanded ResDayAllocationTABLE", each ([ResDayAllocation] = null))
in
    #"Filtered UNMACTHED PERIODS REMOVED";

// Query: MultishiftAvailDayPeriod-Index
// Purpose: Rank available periods by capacity-to-demand ratio with explicit zero-demand protection.
shared #"MultishiftAvailDayPeriod-Index" = let
    Source = MultiShiftAvilabilityDay,
    #"Merged Queries" = Table.NestedJoin(Source, {"Resource", "AvailablePeriod"}, ResDayPeriodAvailabilityTABLE, {"Resource", "Period"}, "ResDayPeriodAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ResDayPeriodAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ResDayPeriodAvailabilityTABLE", {"Availability"}, {"Availability"}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded ResDayPeriodAvailabilityTABLE", {"AvailablePeriod"}, PeriodCapacityTABLE, {"Period"}, "PeriodCapacityTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodCapacityTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "PeriodCapacityTABLE", {"Capacity"}, {"Capacity"}),
    #"Merged Queries2" = Table.NestedJoin(#"Expanded PeriodCapacityTABLE", {"AvailablePeriod"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
    #"Expanded PeriodDemandTABLE" = Table.ExpandTableColumn(#"Merged Queries2", "PeriodDemandTABLE", {"D"}, {"D"}),
    // Preserve the established zero fallback; the priority semantics of excluding zero-demand periods remain a separate business decision.
    #"Inserted C/D" = Table.AddColumn(#"Expanded PeriodDemandTABLE", "C/D", each if [D] = null or [D] = 0 then 0 else [Capacity] / [D], type number),
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
    #"Merged Queries" = Table.NestedJoin(Source, {"Day"}, #"PeriodShiftDay B", {"Day"}, "PeriodShiftDay", JoinKind.LeftOuter),
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
    #"Merged Queries1" = Table.NestedJoin(Source, {"Resource", "Day"}, #"MultiDayPeriod-AvailableAllocatedMatches", {"Resource", "Day"}, "MultiDayPeriod-AvailableAllocatedMatches", JoinKind.LeftOuter),
    #"Expanded MultiDayPeriod-AvailableAllocatedMatches" = Table.ExpandTableColumn(#"Merged Queries1", "MultiDayPeriod-AvailableAllocatedMatches", {"AvailablePeriod"}, {"AvailablePeriod"})
in
    #"Expanded MultiDayPeriod-AvailableAllocatedMatches";

shared MultiPeriodPriority = let
    Source = MultiShiftDayPeriodMatches[[AvailablePeriod]],
    #"Removed Duplicates" = Table.Distinct(Source),
    #"Merged Queries" = Table.NestedJoin(#"Removed Duplicates", {"AvailablePeriod"}, #"PeriodDemandTABLE !!", {"Period"}, "PeriodDemandTABLE", JoinKind.LeftOuter),
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

[ Description = "BUFFER" ]
shared MultiResPeriodPriority = let
    Source = Table.NestedJoin(MultiShiftDayPeriodMatches, {"AvailablePeriod"}, MultiPeriodPriority, {"AvailablePeriod"}, "MultiPeriodPriority", JoinKind.LeftOuter),
    #"Expanded MultiPeriodPriority" = Table.ExpandTableColumn(Source, "MultiPeriodPriority", {"PeriodPriority"}, {"PeriodPriority"}),
    BUFFER = Table.Buffer(#"Expanded MultiPeriodPriority")
in
    BUFFER;

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
    #"Filtered Rows" = Table.SelectRows(Source, each ([NoAllocationKeep] = true)),
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
    #"Grouped Rows" = Table.Group(Source, {"Period"}, {{"Capacity", each List.Sum([Availability]), type number}}),
    BUFFER = Table.Buffer(#"Grouped Rows")
in
    BUFFER;

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

shared ValidAllocation = 0.7 meta [IsParameterQuery=true, Type="Number", IsParameterQueryRequired=true];

shared #"PeriodC'SUM" = let
    Source = #"PeriodC'",
    #"PeriodC'1" = Source[#"PeriodC'"],
    #"Calculated Sum" = List.Sum(#"PeriodC'1")
in
    #"Calculated Sum";

// Query: RolePathTABLE
// Purpose: Resolve this workbook's folder and dynamic role through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl (one-row FilePath table or single named cell) and the public CentriSyncPaths table.
// Output: Existing Variable Name / Value rows used by RolePath, FilePath and Role.
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
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "CapacityDistrib(A.1)-shifts.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Use =CELL(""filename"",A1) in its input cell, then save and recalculate CapacityDistrib(A.1)-shifts.xlsx.",
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

// Query: IMPORT Settings Data
// Purpose: Open Settings Data.xlsx once for all Settings table extractions.
shared #"IMPORT Settings Data" = let
    SourceBinary = Binary.Buffer(File.Contents(FilePath & "\2. Calculations\Settings Data.xlsx")),
    Navigation = Excel.Workbook(SourceBinary, null, true),
    BufferedNavigation = Table.Buffer(Navigation)
in
    BufferedNavigation;

// Query: IMPORTSource Settings Data
// Purpose: Preserve the existing Settings import interface.
shared #"IMPORTSource Settings Data" = #"IMPORT Settings Data";

// Query: EXTRACT PermutationDimensions
// Purpose: Navigate to the Settings permutation table without applying business transformations.
shared #"EXTRACT PermutationDimensions" = let
    Source = #"IMPORT Settings Data"{[Item="PermutationDimensions", Kind="Table"]}[Data]
in
    Source;

// Query: PermutationDimensions Prepare
// Purpose: Apply the dynamic role scope and required Settings types after extraction.
shared #"PermutationDimensions Prepare" = let
    Source = #"EXTRACT PermutationDimensions",
    RequiredColumns = {"Date", "Day", "Shifts", "Period", "RolesList"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(Source)),
    ValidatedTable = if List.IsEmpty(MissingColumns) then Source
        else error Error.Record("A.1 Settings import", "PermutationDimensions is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(ValidatedTable, {{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}}),
    FilteredRole = Table.SelectRows(Typed, each [RolesList] = Role)
in
    FilteredRole;

// Query: EXTRACT Date_From
// Purpose: Validate and return the single DateFrom setting used to scope allocations.
shared #"EXTRACT Date_From" = let
    Source = #"IMPORT Settings Data"{[Item="DateFrom", Kind="Table"]}[Data],
    RequiredColumn = if Table.HasColumns(Source, {"DateFrom"}) then Source
        else error "DateFrom must contain a DateFrom column.",
    SingleRow = if Table.RowCount(RequiredColumn) = 1 then RequiredColumn
        else error "DateFrom must contain exactly one data row.",
    Typed = Table.TransformColumnTypes(SingleRow, {{"DateFrom", type date}}),
    Value = Typed{0}[DateFrom]
in
    Value;

// Query: EXTRACT MaxAvailability
// Purpose: Preserve the legacy Settings maximum interface; employee Effective Shift Cap is canonical for A.1.
// Notes: Retained for compatibility and no longer used by the cap calculation.
shared #"EXTRACT MaxAvailability" = let
    Source = #"IMPORT Settings Data"{[Item="MaxAvailability", Kind="Table"]}[Data],
    RequiredColumn = if Table.HasColumns(Source, {"MaxAvailability"}) then Source
        else error "MaxAvailability must contain a MaxAvailability column.",
    SingleRow = if Table.RowCount(RequiredColumn) = 1 then RequiredColumn
        else error "MaxAvailability must contain exactly one data row.",
    Typed = Table.TransformColumnTypes(SingleRow, {{"MaxAvailability", Int64.Type}}),
    Value = Typed{0}[MaxAvailability]
in
    Value;