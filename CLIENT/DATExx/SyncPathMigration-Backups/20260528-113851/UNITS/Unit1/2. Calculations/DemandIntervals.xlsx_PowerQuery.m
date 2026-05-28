// Power Query from: DemandIntervals.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\DemandIntervals.xlsx
// Extracted: 2026-05-18T06:14:19.082Z

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
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Filtered Rows", {{"Value", each Text.BeforeDelimiter(_, "\", {0, RelativePosition.FromEnd}), type text}}),
    Value = #"Extracted Text Before Delimiter"{0}[Value]
in
    Value;

shared #"IMPORT Table_ShiftDemandHRS" = let

    
    Source = Excel.Workbook(File.Contents(FilePath &  "\1. Input\2-DemandExtract.xlsx"), null, true),
    ShiftUnitDemandHRS_Sheet = Source{[Item="ShiftUnitDemandHRS",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(ShiftUnitDemandHRS_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Day", Int64.Type}, {"Shift", type text}, {"Period", Int64.Type}, {"Role", type text}, {"StartTime", type datetime}, {"EndTime", type datetime}, {"Unit", type any}, {"Facility", type any}, {"DemandFTE", type number}, {"DemandHRS", type number}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"DurationOfShifts", "ShiftDuration"}})
in
    #"Renamed Columns";

shared Table_IntervalIDList = let

    Source = Excel.Workbook(File.Contents(FilePath & "\2. Calculations\Intervals.xlsx"), null, true),
    RoleIntervalID_Table = Source{[Item="RoleIntervalID",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RoleIntervalID_Table,{{"DateTime", type datetime}, {"DayDate", type date}, {"IntervalID", Int64.Type}})
in
    #"Changed Type";

shared Table_Intervals = let
    
    
    
    Source = Excel.Workbook(File.Contents(FilePath &  "\2. Calculations\Intervals.xlsx"), null, true),
    Table_Intervals_Table = Source{[Item="Table_Intervals",Kind="Table"]}[Data],
    #"Sorted Rows" = Table.Sort(Table_Intervals_Table,{{"StartInterval", Order.Ascending}})
in
    #"Sorted Rows";

shared IntervalIDs = let
    Source = Table_IntervalIDList
in
    Source;

shared ShiftINTERVALS = let
    Source = #"IMPORT Table_ShiftDemandHRS",
    #"Grouped Rows" = Table.Group(Source, {"Shift", "Role", "Facility", "StartTime", "EndTime", "Date"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Role", Order.Ascending}, {"StartTime", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Count"}),
    #"Merged DEMAND+STARTINTERVALID" = Table.NestedJoin(#"Removed Columns1", {"StartTime"}, IntervalIDs, {"DateTime"}, "IntervalIDs", JoinKind.LeftOuter),
    #"Expanded IntervalIDs" = Table.ExpandTableColumn(#"Merged DEMAND+STARTINTERVALID", "IntervalIDs", {"DateTime", "DayDate", "IntervalID"}, {"DateTime", "DayDate", "IntervalID"}),
    #"Sorted Rows3" = Table.Sort(#"Expanded IntervalIDs",{{"IntervalID", Order.Ascending}}),
    #"Merged DEMAND+ENDINTERVAL" = Table.NestedJoin(#"Sorted Rows3", {"EndTime"}, IntervalIDs, {"DateTime"}, "IntervalIDs", JoinKind.LeftOuter),
    #"Expanded IntervalIDs1" = Table.ExpandTableColumn(#"Merged DEMAND+ENDINTERVAL", "IntervalIDs", {"DateTime", "IntervalID"}, {"DateTime.1", "IntervalID.1"}),
    #"Sorted Rows4" = Table.Sort(#"Expanded IntervalIDs1",{{"DateTime.1", Order.Ascending}}),
    #"Inserted RANGE" = Table.AddColumn(#"Sorted Rows4", "Range", each if [IntervalID.1] - [IntervalID] <> 0
then [IntervalID.1] - [IntervalID] -1
else 0),
    #"Changed Type" = Table.TransformColumnTypes(#"Inserted RANGE",{{"Range", Int64.Type}}),
    #"Sorted Rows2" = Table.Sort(#"Changed Type",{{"Role", Order.Ascending}, {"StartTime", Order.Ascending}}),
    #"Added INTERVALS" = Table.AddColumn(#"Sorted Rows2", "IntervalListx", each if [IntervalID] = null
then null 
else 
List.Numbers([IntervalID],[Range]+1)),
    #"Expanded IntervalListx" = Table.ExpandListColumn(#"Added INTERVALS", "IntervalListx"),
    #"Removed Duplicates" = Table.Distinct(#"Expanded IntervalListx"),
    #"Removed Columns" = Table.RemoveColumns(#"Removed Duplicates",{"IntervalID", "IntervalID.1", "Range"}),
    #"Sorted Rows1" = Table.Sort(#"Removed Columns",{{"Role", Order.Ascending}, {"IntervalListx", Order.Ascending}})
in
    #"Sorted Rows1";

[ Description = "Modified for ANACC demand#(lf)Not master roster demand" ]
shared ShiftDemandUnitINTERVAL = let
    Source = ShiftINTERVALS,
    #"Merged INTERVALS" = Table.NestedJoin(Source, {"IntervalListx", "Role"}, Table_Intervals, {"IntervalID", "Role"}, "Table_Intervals", JoinKind.LeftOuter),
    #"Expanded Table_Intervals" = Table.ExpandTableColumn(#"Merged INTERVALS", "Table_Intervals", {"StartInterval", "EndInterval", "Duration", "ShiftPeriod"}, {"StartInterval", "EndInterval", "Duration", "ShiftPeriod"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded Table_Intervals",{{"StartInterval", type datetime}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type1",{{"Role", Order.Ascending}, {"IntervalListx", Order.Ascending}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Sorted Rows1",{{"EndInterval", type datetime}}),
    #"Merged SHIFTUNITDEMAND" = Table.NestedJoin(#"Changed Type2", {"Facility", "Role", "StartTime", "EndTime"}, #"IMPORT Table_ShiftDemandHRS", {"Facility", "Role", "StartTime", "EndTime"}, "Table_ShiftDemandHRS", JoinKind.LeftOuter),
    #"Expanded Table_ShiftDemandHRS" = Table.ExpandTableColumn(#"Merged SHIFTUNITDEMAND", "Table_ShiftDemandHRS", {"Unit", "DemandFTE", "DemandHRS", "ShiftDuration"}, {"Unit", "DemandFTE", "DemandHRS", "ShiftDuration"}),
    EFFECTIVEDURATION = Table.AddColumn(#"Expanded Table_ShiftDemandHRS", "EffectiveDuration", each if [ShiftDuration] < MealBreakStart
then [ShiftDuration]
else [ShiftDuration] - 0.5),
    INTEFFECTIVERATIO = Table.AddColumn(EFFECTIVEDURATION, "IntervalEffectiveRatio", each [EffectiveDuration]/[ShiftDuration]),
    //not applicable for ANACC demand - only master roster 
    //#"Added EFFECTIVEDEMAND !!" = Table.AddColumn(INTEFFECTIVERATIO, "EffectiveIntervalAttendance", each [DemandFTE]*[IntervalEffectiveRatio]),
    #"Added EFFECTIVEDEMAND !!" = Table.AddColumn(INTEFFECTIVERATIO, "EffectiveIntervalAttendance", each [DemandFTE]),
    //*[IntervalEffectiveRatio])"
    #"Added DEMAND EFFORT" = Table.AddColumn(#"Added EFFECTIVEDEMAND !!", "DemandEffort", each [Duration] * [EffectiveIntervalAttendance]*24)
in
    #"Added DEMAND EFFORT";

shared ShiftDemandINTERVALLIST = let
    Source = ShiftDemandUnitINTERVAL,
    #"Unpivoted Only Selected Columns" = Table.Unpivot(Source, {"StartInterval", "EndInterval"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Only Selected Columns",{{"Attribute", "PointType"}, {"Value", "TimeDate"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"TimeDate", type datetime}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each true),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"Role", Order.Ascending}, {"IntervalListx", Order.Ascending}, {"EndTime", Order.Ascending},  {"PointType", Order.Descending}})
in
    #"Sorted Rows";

shared Meals = let
    Source = Excel.Workbook(File.Contents(FilePath & "\2. Calculations\Settings Data.xlsx"), null, true),
    Meals_Table = Source{[Item="Meals",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Meals_Table,{{"MealBreakTime", type number}, {"MealBreakstart", Int64.Type}})
in
    #"Changed Type";

shared MealBreakStart = let
    Source = Meals,
    MealBreakstart = Source{0}[MealBreakstart]
in
    MealBreakstart;

shared MealBreakTime = let
    Source = Meals,
    MealBreakTime1 = Source{0}[MealBreakTime]
in
    MealBreakTime1;

shared Table_ShiftDemandHRSCheck = let
    Source = #"IMPORT Table_ShiftDemandHRS",
    #"Grouped Rows" = Table.Group(Source, {"Facility", "Role"}, {{"DemandHrsRoster", each List.Sum([DemandHRS]), type nullable number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "DemandHrsWeek", each [DemandHrsRoster]/2),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "CareDemandHrsPer WeekApprox", each [DemandHrsWeek] *.936)
in
    #"Added Custom1";

shared ShiftDemandUnitINTERVALCheck = let
    Source = ShiftDemandUnitINTERVAL,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Facility"}, {{"DemandHrsPerRoster", each List.Sum([DemandEffort]), type number}}),
    #"Added Custom" = Table.AddColumn(#"Grouped Rows", "DemandHrsPerWeek", each [DemandHrsPerRoster]/2)
in
    #"Added Custom";