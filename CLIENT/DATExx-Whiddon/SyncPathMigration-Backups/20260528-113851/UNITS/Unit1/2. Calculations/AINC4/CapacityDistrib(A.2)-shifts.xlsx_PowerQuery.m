// Power Query from: CapacityDistrib(A.2)-shifts.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\AINC4\CapacityDistrib(A.2)-shifts.xlsx
// Extracted: 2026-05-21T00:53:38.384Z

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
    Custom1 = Table.Buffer(TABLE)
in
    Custom1;

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