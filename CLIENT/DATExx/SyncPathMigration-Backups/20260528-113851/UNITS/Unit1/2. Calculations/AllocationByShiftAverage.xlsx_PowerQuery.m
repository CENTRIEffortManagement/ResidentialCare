// Power Query from: AllocationByShiftAverage.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\AllocationByShiftAverage.xlsx
// Extracted: 2026-05-18T06:14:07.594Z

section Section1;

shared UnitL1PathTABLE = // Version 25.00 250107   
// Updated 250107

let

    
    // PART A - Define FilePathUrl and Replaced Value
    FilePathUrl = 
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePAthUrl"]}[Content],
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
            {"FileName", FileName}
            
        }
    ),
    BUFFER = Table.Buffer(TABLE)
in
    BUFFER;

shared FilePath = let
    Source = UnitL1PathTABLE,
    Value = Source{1}[Value]
in
    Value;

shared #"IMPORT Shifts" = let

    Source = Excel.Workbook(File.Contents(FilePath  &  "\Shifts.xlsx"), null, true),
    BUFFER = Table.Buffer(Source)
in
    BUFFER;

shared #"EXTRACT NAMES_INTERVALSLIST" = let

    Source = #"IMPORT Shifts",
    NAMES_INTERVALSLIST_Table = Source{[Item="NAMES_INTERVALSLIST",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(NAMES_INTERVALSLIST_Table,{{"DayDate", type date}, {"Name", type text}, {"TimeType", type text}, {"AllocatedIntervals", Int64.Type}, {"IntervalID.1", Int64.Type}, {"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type datetime}, {"IntervalDurationTemp", type number}, {"Double Shift", type any}, {"RealDuration", type number}, {"Effective Duration", type number}, {"IntervalAssociatedShift", type text}, {"Attribute", type text}, {"ShiftType", type text}, {"ShiftDate", type date}, {"Role1", type text}, {"StaffCount", Int64.Type}})
in
    #"Changed Type";

shared #"EXTRACT ShiftsDayTime" = let
   
    Source = #"IMPORT Shifts",
    ShiftsDayTime_Table = Source{[Item="ShiftsDayTime",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftsDayTime_Table,{{"DateTime", type datetime}, {"Shift", type text}, {"TimeType", type text}})
in
    #"Changed Type";

shared #"IMPORT ShiftDuration" = let
     
    Source = Excel.Workbook(File.Contents(FilePath&"\Settings Data.xlsx"), null, true),
    ShiftDuration_Table = Source{[Item="ShiftDuration",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDuration_Table,{{"StdWeekDays", Int64.Type}, {"ShiftDuration", type number}}),
    ShiftDuration1 = #"Changed Type"{0}[ShiftDuration]
in
    ShiftDuration1;

shared #"IMPORT ShiftPeriod" = let

    
    Source = Excel.Workbook(File.Contents(FilePath&"\Settings Data.xlsx"), null, true),
    ShiftPeriod_Table = Source{[Item="ShiftPeriod",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftPeriod_Table,{{"ShiftPeriod", type text}, {"StartTime", type number}, {"StartDay", type number}})
in
    #"Changed Type";

shared AMP = let
    Source = #"IMPORT ShiftPeriod",
    StartDay = Source{0}[StartDay]
in
    StartDay;

shared Roles = let
    Source = #"EXTRACT NAMES_INTERVALSLIST",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Role"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns")
in
    #"Removed Duplicates";

shared RoleShiftsDayTime = let
    Source = #"EXTRACT ShiftsDayTime",
    #"Removed Columns" = Table.RemoveColumns(Source,{"TimeType"}),
    #"Added ROLETABLE" = Table.AddColumn(#"Removed Columns", "Roles", each Roles),
    #"Expanded Roles" = Table.ExpandTableColumn(#"Added ROLETABLE", "Roles", {"Role"}, {"Roles.Role"})
in
    #"Expanded Roles";

shared ResourceShiftAllocation = let
    Source = ResourceIntervalAllocation,
    #"Grouped Rows" = Table.Group(Source, {"ShiftDate", "ShiftPeriod", "IntervalAssociatedShift", "Name", "Role", "TimeDate", "Attribute"}, {{"ResShiftEffort", each List.Sum([ResEffectiveIntervalEffort]), type number}, {"ResShiftEffectiveRatio", each List.Average([#"ResEffectiveRatio-FTE"]), type number}}),
    #"Added RESSHIFTFTE" = Table.AddColumn(#"Grouped Rows", "ResShiftFTE", each [ResShiftEffort]*24/#"IMPORT ShiftDuration"),
    #"Filtered Rows" = Table.SelectRows(#"Added RESSHIFTFTE", each ([Attribute] = "IntervalStart"))
in
    #"Filtered Rows";

shared #"ResourceShiftAllocation - less Intervals" = let
    Source = ResourceShiftAllocation,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"TimeDate", type date}}),
    #"Grouped Rows" = Table.Group(#"Changed Type", {"ShiftPeriod", "Name", "Role", "TimeDate"}, {{"ResShiftFTE", each List.Sum([ResShiftFTE]), type number}})
in
    #"Grouped Rows";

[ Description = "BUFFER" ]
shared ResourceIntervalAllocation = let
    Source = #"EXTRACT NAMES_INTERVALSLIST",
    #"Filtered Rows" = Table.SelectRows(Source, each ([RealDuration] <> 0)),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"IntervalDurationTemp", "IntervalDuration"}}),
    RESEFFECTIVERATIO = Table.AddColumn(#"Renamed Columns", "ResEffectiveRatio-FTE", each if [StaffCount] = 0 then 0
else [Effective Duration] / [RealDuration]),
    EFFECTIVEINTERVALEFFORT = Table.AddColumn(RESEFFECTIVERATIO, "ResEffectiveIntervalEffort", each [#"ResEffectiveRatio-FTE"]*[IntervalDuration]),
    #"Removed Columns" = Table.RemoveColumns(EFFECTIVEINTERVALEFFORT,{ "TimeType"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"TimeDate", type datetime}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Changed Type",{{"ShiftDate", type date}}),
    BUFFER = Table.Buffer(#"Changed Type1")
in
    BUFFER;

shared RoleShiftAllocation = let
    Source = ResourceIntervalAllocation,
    #"Filtered Rows" = Table.SelectRows(Source, each ([StaffCount] = 1)),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"ShiftDate", "ShiftPeriod", "Role", "Attribute"}, {{"RoleShiftEffort", each List.Sum([ResEffectiveIntervalEffort]), type number}}),
    #"Filtered Rows1" = Table.SelectRows(#"Grouped Rows", each ([Attribute] = "IntervalStart")),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows1", {"ShiftPeriod", "Role"}, #"IMPORT ShiftPeriod", {"ShiftPeriod", "Role"}, "ShiftPeriod.1", JoinKind.LeftOuter),
    #"Expanded ShiftPeriod.1" = Table.ExpandTableColumn(#"Merged Queries", "ShiftPeriod.1", {"DurationOfShifts"}, {"DurationOfShifts"}),
    ROLESHIFTFTE = Table.AddColumn(#"Expanded ShiftPeriod.1", "RoleShiftFTE", each [RoleShiftEffort]*24/[DurationOfShifts]),
    #"Sorted Rows2" = Table.Sort(ROLESHIFTFTE,{{"Role", Order.Ascending}, {"ShiftDate", Order.Ascending}})
in
    #"Sorted Rows2";

shared RoleShiftIntervalAllocation = let
    Source = ResourceIntervalAllocation,
    #"Grouped Rows" = Table.Group(Source, {"ShiftDate", "ShiftPeriod", "Role", "Attribute"}, {{"RoleShiftEffort", each List.Sum([ResEffectiveIntervalEffort]), type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Grouped Rows", {"ShiftPeriod"}, #"IMPORT ShiftPeriod", {"ShiftPeriod"}, "ShiftPeriod.1", JoinKind.LeftOuter),
    #"Expanded ShiftPeriod.1" = Table.ExpandTableColumn(#"Merged Queries", "ShiftPeriod.1", {"DurationOfShifts"}, {"DurationOfShifts"}),
    ROLESHIFTFTE = Table.AddColumn(#"Expanded ShiftPeriod.1", "RoleShiftFTE", each [RoleShiftEffort]*24/[DurationOfShifts]),
    #"Sorted Rows" = Table.Sort(ROLESHIFTFTE,{{"ShiftDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    #"Sorted Rows";

shared RoleShiftEffortFTE = let
    Source = ResourceIntervalAllocation,
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayDate", type date}})
in
    #"Changed Type";