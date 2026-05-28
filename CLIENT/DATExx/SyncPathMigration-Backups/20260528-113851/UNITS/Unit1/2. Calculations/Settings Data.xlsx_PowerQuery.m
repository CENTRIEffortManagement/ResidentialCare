// Power Query from: Settings Data.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Settings Data.xlsx
// Extracted: 2026-05-18T06:14:28.949Z

section Section1;

shared #"Dates Listed" = let
    Source = DateRoleShiftAllocation,
    #"Removed Columns" = Table.RemoveColumns(Source,{"Role", "Shift", "Allocation"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Sorted Rows" = Table.Sort(#"Removed Duplicates",{{"Date", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared #"Min Date" = let
    Source = #"Dates Listed",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Date] >= DateFrom)),
    Custom1 = Table.Min(#"Filtered Rows","Date"),
    Date = Custom1[Date]
in
    Date;

shared #"Max Date" = let
    Source = #"Dates Listed",
    Custom1 = Table.Max(Source,"Date")
in
    Custom1;

shared DateList = let
    Source = #"Dates Listed",
    #"Changed Type4" = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Custom1 = #"Changed Type4"[Date],
    Custom2 = List.Max(#"Changed Type4" [Date]),
    #"Converted to Table" = #table(1, {{Custom2}}),
    #"Renamed Columns" = Table.RenameColumns(#"Converted to Table",{{"Column1", "Max"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Min", each #"Min Date"),
    #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom1",{{"Max", type date}, {"Min", type date}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type3",{{"Max", Int64.Type}, {"Min", Int64.Type}}),
    #"Added Custom3" = Table.AddColumn(#"Changed Type2", "Custom", each {[Min]..[Max]}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom3",{{"Max", type date}, {"Min", type date}}),
    #"Expanded Custom" = Table.ExpandListColumn(#"Changed Type1", "Custom"),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Custom",{"Max", "Min"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Custom", type date}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Changed Type",{{"Custom", "Date"}}),
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns1", "Day", 1, 1, Int64.Type)
in
    #"Added Index";

shared PermutationDimensions = let
    Source = #"DateList",
    #"Added SHIFTS" = Table.AddColumn(Source, "Shifts", each Shifts),
    #"Expanded Shifts" = Table.ExpandListColumn(#"Added SHIFTS", "Shifts"),
    #"Added Index" = Table.AddIndexColumn(#"Expanded Shifts", "Period", 1, 1, Int64.Type),
    #"ADD ROLES" = Table.AddColumn(#"Added Index", "RolesList", each Roles),
    #"Expanded RolesList" = Table.ExpandListColumn(#"ADD ROLES", "RolesList"),
    #"Sorted Rows" = Table.Sort(#"Expanded RolesList",{{"Date", Order.Ascending}, {"Period", Order.Ascending}, {"RolesList", Order.Ascending}})
in
    #"Sorted Rows";

shared DateNameRoleShiftAllocation = let
    Source = AllocationExtracted,
    #"Changed Type1" = Table.TransformColumnTypes(Source,{{"Start", type number}, {"End", type number}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "Custom", each (if [End]<[Start]
then (1-[Start])+[End]
else [End]-[Start])*24),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom", each ([End] <> null)),
    #"Added Custom2" = Table.AddColumn(#"Filtered Rows", "TimeWorked", each if [Custom] >6 then [Custom] - .5
else [Custom]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom2", "Shift", each if [Start] < 0.604 then "AM" else if [Start] > 0.89 then "NIGHT" else "PM"),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom1",{ "Start", "End", "Custom"})
in
    #"Removed Columns1";

shared DateRoleShiftAllocation = let
    Source = DateNameRoleShiftAllocation,
    #"Grouped Rows" = Table.Group(Source, {"Date", "Role", "Shift"}, {{"Allocation", each List.Sum([TimeWorked]), type number}})
in
    #"Grouped Rows";

shared Folder = let
    Source = #"FilePath-Facility"
in
    Source;

shared AllocationExtracted = let
    Source1 = Folder,
    Source = Excel.Workbook(File.Contents(Folder & "\1. Input\1-AllocationExtracted.xlsx"), null, true),
    AllocationExtracted_Sheet = Source{[Item="AllocationExtracted",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(AllocationExtracted_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Date", type date}, {"Start", type time}, {"End", type time}, {"Break Time", type time}, {"Hours", type number}, {"Name", type text}, {"Code", Int64.Type}})
in
    #"Changed Type";

shared Roles = let
    Source = Excel.CurrentWorkbook(){[Name="Roles"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"ROLES", type text}}),
    ROLES = #"Changed Type"[ROLES]
in
    ROLES;

shared Unit = let
    Source = Excel.CurrentWorkbook(){[Name="Unit"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"UNITS", type text}}),
    WARDS = #"Changed Type"[UNITS]
in
    WARDS;

shared Shifts = let
    Source = Excel.CurrentWorkbook(){[Name="Shifts"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Shifts", type text}}),
    #"Added Index1" = Table.AddIndexColumn(#"Changed Type", "Index.1", 1, 1, Int64.Type),
    Shifts1 = #"Added Index1"[Shifts]
in
    Shifts1;

shared Query1 = let
    Source = Folder
in
    Source;

shared DateFrom = let
    Source = Excel.CurrentWorkbook(){[Name="DateFrom"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DateFrom", type date}}),
    DateFrom1 = #"Changed Type"{0}[DateFrom]
in
    DateFrom1;

shared IMPORTRootPath = let
    Source = Table.FromColumns({Lines.FromBinary(File.Contents("C:\Users\Alex\Centri\4. Production - Documents\WFEffectiveness\4.1.1 AGED CARE\RootPath.txt"), null, null, 1252)}),
    Column1 = Source{0}[Column1]
in
    Column1;

[ Description = "BUFFER" ]
shared #"INPUT FilePath B" = let
    Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
    #"Renamed Columns" = Table.RenameColumns(Source,{{"FilePathUrl", "String"}})
in
    #"Renamed Columns";

shared Client = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 7), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared Date = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 8), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared Facility = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter1" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", 10), type text}}),
    #"Extracted Text Before Delimiter1" = Table.TransformColumns(#"Extracted Text After Delimiter1", {{"String", each Text.BeforeDelimiter(_, "/"), type text}}),
    String = #"Extracted Text Before Delimiter1"{0}[String]
in
    String;

shared #"FilePath-Facility" = let
    Source = IMPORTRootPath&"\"&Client&"\"&Date&"\"&"FACILITIES"&"\"&Facility
in
    Source;

shared FileName = let
    Source = #"INPUT FilePath B",
    #"Extracted Text After Delimiter" = Table.TransformColumns(Source, {{"String", each Text.AfterDelimiter(_, "/", {0, RelativePosition.FromEnd}), type text}}),
    #"Extracted Text Before Delimiter" = Table.TransformColumns(#"Extracted Text After Delimiter", {{"String", each Text.BeforeDelimiter(_, "]"), type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Extracted Text Before Delimiter","[","",Replacer.ReplaceText,{"String"}),
    String = #"Replaced Value"{0}[String]
in
    String;

shared #"File Path Data" = let
   

    // Create the table with variable names and their corresponding values
    Source = #table(
        {"Variable Name", "Value"},
        {
            {"Root Path", #"IMPORTRootPath"},
            {"FilePath", #"FilePath-Facility"},
            {"Client", Client},
            {"Date", Date},
            {"Facility", Facility},
            {"FileName", FileName}
            
            
            
            
            
            
        }
    )
in
    Source;

shared #"PermutationDimensions (2)" = let
    Source = #"DateList",
    #"Added SHIFTS" = Table.AddColumn(Source, "Shifts", each Shifts),
    #"Expanded Shifts" = Table.ExpandListColumn(#"Added SHIFTS", "Shifts"),
    #"Added Index" = Table.AddIndexColumn(#"Expanded Shifts", "Period", 1, 1, Int64.Type),
    #"ADD ROLES" = Table.AddColumn(#"Added Index", "RolesList", each Roles),
    #"Expanded RolesList" = Table.ExpandListColumn(#"ADD ROLES", "RolesList"),
    #"Sorted Rows" = Table.Sort(#"Expanded RolesList",{{"Date", Order.Ascending}, {"Period", Order.Ascending}, {"RolesList", Order.Ascending}})
in
    #"Sorted Rows";

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