// Power Query from: CapacityDistrib(A.2)-shifts.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\RN\CapacityDistrib(A.2)-shifts.xlsx
// Extracted: 2026-05-21T00:53:54.000Z

section Section1;

[ Description = "BUFFER" ]
shared ResCapPrioritisedTABLE = let
    Source = #"Period Running Total",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Resource", "Period", "PRCell", "Reduction", "ResAv-C'", "ClusterAllocation", "NWDPriority", "NWDType", "MaxPeriodReduction", "PeriodRunningTotal", "ReducePeriod"}),
    #"Sorted RES,NWDPRI,MAXPRT,PRT" = Table.Sort(#"Removed Other Columns",{{"Resource", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"MaxPeriodReduction", Order.Descending}, {"PeriodRunningTotal", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted RES,NWDPRI,MAXPRT,PRT", "ResIndex", 2, 1, Int64.Type),
    BUFFER = Table.Buffer(#"Added Index")
in
    BUFFER;

shared ResCapIndexLimits = let
    Source = ResCapPrioritisedTABLE,
    #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1, Int64.Type),
    #"Renamed Columns" = Table.RenameColumns(#"Added Index",{{"Index", "ResIndexX"}}),
    #"Grouped Rows" = Table.Group(#"Renamed Columns", {"Resource"}, {{"StartResIndex", each List.Min([ResIndexX]), type number}, {"ResUnassignedAvailable", each List.Max([#"ResAv-C'"]), type nullable number}, {"ResAvailNumber", each Table.RowCount(_), Int64.Type}})
in
    #"Grouped Rows";

[ Description = "BUFFER. Cap rs. availability based on roster" ]
shared #"Resource Running total" = let
    Source = Table.NestedJoin(ResCapPrioritisedTABLE, {"Resource"}, ResCapIndexLimits, {"Resource"}, "ResCapLimits", JoinKind.LeftOuter),
    #"Expanded REsLimits" = Table.ExpandTableColumn(Source, "ResCapLimits", {"StartResIndex", "ResAvailNumber"}, {"StartResIndex", "ResAvailNumber"}),
    #"Sorted Rows" = Table.Sort(#"Expanded REsLimits",{{"ResIndex", Order.Ascending}, {"ClusterAllocation", Order.Ascending}, {"PeriodRunningTotal", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction.1", each [ResIndex]- [StartResIndex]),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Subtraction",{{"Subtraction.1", Int64.Type}}),
    LIST.BUFFER = List.Buffer(#"Changed Type"[Reduction]),
    #"Added RUNNINGRESTOTAL" = Table.AddColumn(#"Changed Type", "ResRunningTotal", each List.Sum(List.Range(Source [Reduction], 
[StartResIndex]-1

, [Subtraction.1]))),
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Added RUNNINGRESTOTAL", "MaxResReduction", each List.Min({ [#"ResAv-C'"]})),
    #"Added KEEP" = Table.AddColumn(#"Inserted MAXRESTOTAL", "KeepRunningTotal", each if [ResRunningTotal] > [MaxResReduction] then "Remove for Res" else if [PeriodRunningTotal] > [MaxPeriodReduction] then "Remove for Perod" else if [ResRunningTotal] <= [MaxResReduction] then "Keep" else "Split"),
    #"Filtered KEEP" = Table.SelectRows(#"Added KEEP", each ([KeepRunningTotal] = "Keep")),
    #"Multiplied Column" = Table.TransformColumns(#"Filtered KEEP", {{"Reduction", each _ * -1, type number}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Multiplied Column",{"Resource", "Period", "PRCell", "Reduction", "ResRunningTotal"}),
    BUFFER = Table.Buffer(#"Removed Other Columns")
in
    BUFFER;

shared CheckCap = let
    Source = #"Resource Running total",
    #"Grouped Rows" = Table.Group(Source, {"Resource"}, {{"MAXResRunningTotal", each List.Max([ResRunningTotal]), type number}, {"AvailabilityReduction", each List.Sum([Reduction]), type number}})
in
    #"Grouped Rows";

shared #"ResPeriodCappedAvailability(C#)TABLE" = let
    Source = #"IMPORT ResPeriodAvailabilityTABLE",
    
    #"Merged PREMPTY" = Table.NestedJoin(Source, {"Role", "Resource", "Period"}, #"IMPORT ResourcePeriodTABLE_empty", {"Role", "Resource", "Period"}, "ResourcePeriodTABLE_empty", JoinKind.RightOuter),
    #"Removed Columns1" = Table.RemoveColumns(#"Merged PREMPTY",{"Resource", "Period"}),
    #"Expanded ResourcePeriodTABLE_empty" = Table.ExpandTableColumn(#"Removed Columns1", "ResourcePeriodTABLE_empty", {"Resource", "Period"}, {"Resource", "Period"}),
    
    #"Merged Queries" = Table.NestedJoin(#"Expanded ResourcePeriodTABLE_empty", {"Resource", "Period"}, #"Resource Running total", {"Resource", "Period"}, "ReCapAvailabilityTABLE", JoinKind.LeftOuter),
    #"Expanded ReCapAvailabilityTABLE" = Table.ExpandTableColumn(#"Merged Queries", "ReCapAvailabilityTABLE", {"Reduction"}, {"Reduction"}),
    #"Sorted Rows" = Table.Sort(#"Expanded ReCapAvailabilityTABLE",{{"Resource", Order.Ascending}, {"Period", Order.Ascending}}),
    
    #"Merged Queries1" = Table.NestedJoin(#"Sorted Rows", {"Resource", "Period"}, #"IMPORT ResPeriodWDTABLE", {"Resource", "Period"}, "ResPeriodWDTABLE", JoinKind.FullOuter),
    #"Expanded ResPeriodWDTABLE" = Table.ExpandTableColumn(#"Merged Queries1", "ResPeriodWDTABLE", {"RosteredPeriodStatus"}, {"RosteredPeriodStatus"}),
    #"Replaced Value" = Table.ReplaceValue(#"Expanded ResPeriodWDTABLE",null,0,Replacer.ReplaceValue,{"Reduction"}),
    
    #"Merged Queries2" = Table.NestedJoin(#"Replaced Value", {"Resource", "Period"}, #"IMPORT MultiPeriod - Remove", {"Resource", "AvailablePeriod"}, "MultiDayPeriod-Remove", JoinKind.LeftOuter),
    #"Expanded MultiDayPeriod-Remove" = Table.ExpandTableColumn(#"Merged Queries2", "MultiDayPeriod-Remove", {"NoAllocationKeep"}, {"NoAllocationKeep"}),
    #"Inserted REDUCTION" = Table.AddColumn(#"Expanded MultiDayPeriod-Remove", "AvailabilityCapped", each if (
        [RosteredPeriodStatus] = "RosteredDay" 
        or
        [RosteredPeriodStatus] = "UnavailableAllocatedPeriod" 
        or 
        [Availability] = null
        or 
        [NoAllocationKeep] = false 
        ) 
    then 0
    else 
        [Availability] + [Reduction]),
        #"Removed Columns" = Table.RemoveColumns(#"Inserted REDUCTION",{"Availability", "Reduction", "NoAllocationKeep"}),
    #"Filled Down" = Table.FillDown(#"Removed Columns",{"Role"})
    in
        #"Filled Down";

[ Description = "BUFFER-All cells capped availability" ]
shared #"ResPeriodAvailabilityCapped(C#)TABLE" = let
    Source = #"ResPeriodCappedAvailability(C#)TABLE"
in
    Source;

shared #"ResPeriodAvailabilityCapped(C#)SUM" = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    AvailabilityCapped = Source[AvailabilityCapped],
    #"Calculated Sum" = List.Sum(AvailabilityCapped)
in
    #"Calculated Sum";

shared ResPeriodAvailabilityCappedMATRIX = let
    Source = #"ResPeriodAvailabilityCapped(C#)TABLE",
    #"Removed Columns" = Table.RemoveColumns(Source,{"RosteredPeriodStatus"}),
    #"Pivoted Column" = Table.Pivot(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU"), List.Distinct(Table.TransformColumnTypes(#"Removed Columns", {{"Period", type text}}, "en-AU")[Period]), "Period", "AvailabilityCapped", List.Sum)
in
    #"Pivoted Column";

shared PeriodCapPrioritisedTABLE = let
    Source = #"IMPORT ResPeriodCapPrioritised",
    #"Sorted NWD,C-D,PA-Cn,PERIOD" = Table.Sort(Source,{{"Period", Order.Ascending}, {"NWDPriority", Order.Ascending}, {"ClusterAllocation", Order.Descending}, {"C'-D", Order.Ascending}, {"PeriodA-CNeg", Order.Ascending} }),
    #"Added Index" = Table.AddIndexColumn(#"Sorted NWD,C-D,PA-Cn,PERIOD", "PeriodIndex",  2, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "PeriodIndexX", 1, 1, Int64.Type)
in
    #"Added Index1";

shared PeriodCapIndexLimits = let
    Source = PeriodCapPrioritisedTABLE,
    #"Grouped INDEX+COUNT" = Table.Group(Source, {"Period"}, {{"StartPeriodIndex", each List.Min([PeriodIndexX]), type number}, {"PeriodAvailNumber", each Table.RowCount(Table.Distinct(_)), Int64.Type}, {"PeriodResAv-C'", each List.Max([#"ResAv-C'"]), type nullable number}})
in
    #"Grouped INDEX+COUNT";

[ Description = "BUFFER. Cap rs. availability based on roster" ]
shared #"Period Running Total" = let
    Source = Table.NestedJoin(PeriodCapPrioritisedTABLE, {"Period"}, PeriodCapIndexLimits, {"Period"}, "ResCapLimits", JoinKind.LeftOuter),
    #"Expanded PeriodLimits" = Table.ExpandTableColumn(Source, "ResCapLimits", {"StartPeriodIndex", "PeriodAvailNumber", "PeriodResAv-C'"}, {"StartPeriodIndex", "PeriodAvailNumber", "PeriodResAv-C'"}),
    #"Sorted Rows" = Table.Sort(#"Expanded PeriodLimits",{{"PeriodIndex", Order.Ascending}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Sorted Rows", "Subtraction.1", each [PeriodIndex]-[StartPeriodIndex]),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted Subtraction",{{"Subtraction.1", Int64.Type}}),
    // Min reduction: PeriodA-C, PeriodC-D
    #"Inserted MAXRESTOTAL" = Table.AddColumn(#"Changed Type", "MaxPeriodReduction", each List.Min({-1*[#"PeriodA-CNeg"], [#"C'-D"]})),
    #"BUFFER LIST" = List.Buffer(#"Inserted MAXRESTOTAL"[UnassignedAvail]),
    #"Added RUNNINGRESTOTAL" = Table.AddColumn(#"Inserted MAXRESTOTAL", "PeriodRunningTotal", each List.Sum(List.Range(Source [UnassignedAvail], [StartPeriodIndex]-1

, [Subtraction.1]
))),
    #"Renamed Columns" = Table.RenameColumns(#"Added RUNNINGRESTOTAL",{{"UnassignedAvail", "Reduction"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns", "ReducePeriod", each if [PeriodRunningTotal] <= [MaxPeriodReduction] then "RemovableByPeriod" else "UnremovableByPeriod"),
    BUFFER = Table.Buffer(#"Added Conditional Column")
in
    BUFFER;

shared ResPeriodAvailabilityCappedMATRIXSUM = let
    Source = ResPeriodAvailabilityCappedMATRIX,
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(Source, {"Resource","Role"}, "Attribute", "Value"),
    Value = #"Unpivoted Other Columns"[Value],
    #"Calculated Sum" = List.Sum(Value)
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

shared #"IMPORT ResPeriodCapPrioritised" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodCapPrioritised_Table = Source{[Item="ResPeriodCapPrioritised",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodCapPrioritised_Table,{{"Resource", Int64.Type}, {"Period", Int64.Type}, {"Allocation", type any}, {"PotentialAvailability", type any}, {"PRCell", type text}, {"UnassignedAvail", Int64.Type}, {"ResAv-C'", Int64.Type}, {"PeriodA-CNeg", type number}, {"C'-D", Int64.Type}, {"C'/D", type number}, {"ClusterAllocation", type number}, {"NWDPriority", Int64.Type}, {"NWDType", type text}})
in
    #"Changed Type";

shared #"IMPORT MultiPeriod - Remove" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    MultiDayPeriod_Remove_Table = Source{[Item="MultiDayPeriod_Remove",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(MultiDayPeriod_Remove_Table,{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"C/D", type number}, {"NoAllocationKeep", type logical}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type2",{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"C/D", type number}, {"NoAllocationKeep", type logical}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type",{{"Resource", Int64.Type}, {"AvailablePeriod", Int64.Type}, {"C/D", type number}, {"NoAllocationKeep", type logical}})
in
    #"Changed Type1";

shared #"IMPORT ResPeriodWDTABLE" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodWDTABLE_Table = Source{[Item="ResPeriodWDTABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodWDTABLE_Table,{{"Resource", Int64.Type}, {"Day", Int64.Type}, {"PotentialAvailability", type text}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"RosteredPeriodStatus", type text}})
in
    #"Changed Type1";

shared #"IMPORT ResPeriodAvailabilityTABLE" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}})
in
    #"Changed Type";

shared #"IMPORT ResourcePeriodTABLE_empty" = let
    Source = Excel.Workbook(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResourcePeriodTABLE_empty_Table = Source{[Item="ResourcePeriodTABLE_empty",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResourcePeriodTABLE_empty_Table,{{"Name", type text},  {"Role", type text},     {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Shift", type any}, {"Day", Int64.Type}})
in
    #"Changed Type";
