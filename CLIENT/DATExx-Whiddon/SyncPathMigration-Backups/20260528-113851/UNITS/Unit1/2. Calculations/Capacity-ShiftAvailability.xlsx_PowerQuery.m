// Power Query from: Capacity-ShiftAvailability.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Capacity-ShiftAvailability.xlsx
// Extracted: 2026-05-18T06:14:10.401Z

section Section1;

shared #"IMPORT AvailabilityDayShiftMATRIXRaw" = let
 
    Source = Excel.Workbook(File.Contents(#"FilePath - 1Input" & "\StaffList Availability.xlsx"), null, true),
    Table2_Table = Source{[Item="AvailabilityMatrix",Kind="Table"]}[Data]
in
    Table2_Table;

shared #"IMPORT OrdinaryWorkWk" = let

    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true),
    Table13_Table = Source{[Item="Table13",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table13_Table,{{"OrdinaryHrsWk-Week", Int64.Type}, {"OrdinaryHrsFn", Int64.Type}, {"OrdinaryHrs28d", Int64.Type}}),
    #"OrdinaryHrsWk-Week" = #"Changed Type"{0}[#"OrdinaryHrsWk-Week"]
in
    #"OrdinaryHrsWk-Week";

shared #"IMPORT MultiRolesResRolesProportion" = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 1Input" & "\1-AllocationExtracted.xlsx"), null, true),
    ResRolesProportion_Table = Source{[Item="ResRolesProportion",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResRolesProportion_Table,{{"Name", type text}, {"AINC4", type number}, {"AIN", type number}, {"Total", type number}, {"AINC4HrsAvilPref", type number}})
in
    #"Changed Type";

shared #"IMPORT Table_Min_Date" = let
    Source = Excel.Workbook(File.Contents(#"FilePath - 2Calculations" & "\Settings Data.xlsx"), null, true),
    Table_Min_Date_Table = Source{[Item="Table_Min_Date", Kind="Table"]}[Data],
    #"Min Date" = Table_Min_Date_Table{0}[Min Date]
in
    #"Min Date"
;

shared EffectiveShiftHrs = let
    Source = #"IMPORT OrdinaryWorkWk"/StdRosterDays*2
in
    Source;

[ Description = "Max number of shofts available each week" ]
shared StdRosterDays = 10 meta [IsParameterQuery=true, Type="Number", IsParameterQueryRequired=true];

shared #"AvailabilityDayShift Prepare" = let
    Source = #"IMPORT AvailabilityDayShiftMATRIXRaw",
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Full Name", "Name"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"ID Number"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Removed Columns",{"ID", "Employment Status", "Contracted Weekly Hours"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Removed Columns1", {"Role", "Name"  }, "WeekDayShift", "Available")
in
    #"Unpivoted Other Columns";

shared ResDayShift = let
    Source = #"AvailabilityDayShift Prepare",
    #"Filtered Rows" = Table.SelectRows(Source, each ([WeekDayShift] <> "Column1")),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Filtered Rows", "WeekDayShift", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"Week", "Day", "Shift"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Split Column by Delimiter",{"Available"}),
    #"Extracted First Characters" = Table.TransformColumns(#"Removed Columns1", {{"Shift", each Text.Start(_, 2), type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Extracted First Characters","NI","NIGHT",Replacer.ReplaceText,{"Shift"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","M2","PM",Replacer.ReplaceText,{"Shift"}),
    EFFECTTIVESHIFTHORS = Table.AddColumn(#"Replaced Value1", "EffectiveShiftHrs", each EffectiveShiftHrs),
    #"Merged DAYORDER" = Table.NestedJoin(EFFECTTIVESHIFTHORS, {"Day"}, DayOrder, {"DayShort"}, "DayOrder", JoinKind.LeftOuter),
    #"Expanded DayOrder" = Table.ExpandTableColumn(#"Merged DAYORDER", "DayOrder", {"DayOrder"}, {"DayOrder.1"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded DayOrder",{{"Week", Int64.Type}}),
    #"Filtered Rows1" = Table.SelectRows(#"Changed Type", each ([Role] <> null)),
    #"Inserted ROSTERDATEORDER" = Table.AddColumn(#"Filtered Rows1", "RosterDateOrder", each (([Week]-1) * 7)+[DayOrder.1]),
    #"Sorted Rows" = Table.Sort(#"Inserted ROSTERDATEORDER",{{"Week", Order.Ascending}, {"RosterDateOrder", Order.Ascending}}),
    #"Added DATE" = Table.AddColumn(#"Sorted Rows", "Date", each Date.AddDays(#"IMPORT Table_Min_Date",[RosterDateOrder]-1)),
    #"Removed Columns" = Table.RemoveColumns(#"Added DATE",{"DayOrder.1", "RosterDateOrder"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Name", "NameX"}}),
    #"Merged Queries" = Table.NestedJoin(#"Renamed Columns", {"NameX"}, #"IMPORT MultiRolesResRolesProportion", {"Name"}, "IMPORT MultiRolesResRolesProportion", JoinKind.LeftOuter),
    #"Expanded IMPORT MultiRolesResRolesProportion" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT MultiRolesResRolesProportion", {"Name"}, {"NameY"}),
    #"Inserted Merged Column" = Table.AddColumn(#"Expanded IMPORT MultiRolesResRolesProportion", "Name", each if [NameY] = null then [NameX] 
else 
Text.Combine({[NameY], [Role]}, " (")&")"),
    #"Removed Columns2" = Table.RemoveColumns(#"Inserted Merged Column",{"NameX", "NameY"})
in
    #"Removed Columns2";

shared DayOrder = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45W8s3PU9JRMlSK1YlWCilNBbKNwOzw1BQg2xginlEKZJuA2W5FmUC2KZgdnFgCZJtB2KUgc8yVYmMB", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [DayShort = _t, DayOrder = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DayShort", type text}, {"DayOrder", Int64.Type}})
in
    #"Changed Type";

shared MultiRoles = let
    Source = #"IMPORT MultiRolesResRolesProportion",
    #"Removed Columns1" = Table.RemoveColumns(Source,{"Total"}),
    #"Unpivoted Other Columns1" = Table.UnpivotOtherColumns(#"Removed Columns1", {"Name", "AINC4HrsAvilPref"}, "MultiRole", "Value"),
    #"Removed Columns" = Table.RemoveColumns(#"Unpivoted Other Columns1",{"Value"})
in
    #"Removed Columns";

shared RoleResDayAvailabilityCapped = let
    Source = ResDayShift,
    #"Grouped Rows" = Table.Group(Source, {"Role", "Name"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Added STDDAYSAVAIL" = Table.AddColumn(#"Grouped Rows", "StdRosterDays", each StdRosterDays),
    #"Merged Queries" = Table.NestedJoin(#"Added STDDAYSAVAIL", {"Name"}, MultiRoles, {"Name"}, "IMPORT ResRolesProportion2", JoinKind.LeftOuter),
    #"Expanded IMPORT ResRolesProportion2" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT ResRolesProportion2", {"AINC4HrsAvilPref", "MultiRole"}, {"AINC4HrsAvilPref", "MultiRole"}),
    #"Added ROLESPLIT" = Table.AddColumn(#"Expanded IMPORT ResRolesProportion2", "RoleSplit", each if [MultiRole] = null then null  
else if [MultiRole] = "AIN" and [MultiRole] = [Role] then 1 - [AINC4HrsAvilPref]
else if [MultiRole] = [Role] then [AINC4HrsAvilPref]
else 99),
    #"Filtered ROLESPLIT" = Table.SelectRows(#"Added ROLESPLIT", each ([RoleSplit] <> 99)),
    #"Added SPLITAVAIL" = Table.AddColumn(#"Filtered ROLESPLIT", "AvailabilityCapped", each if [Count] <= [StdRosterDays] then [Count] else [StdRosterDays] * (if [MultiRole] <> null then [RoleSplit] else 1)),
    #"Transformed Columns" = Table.TransformColumns(
        #"Added SPLITAVAIL",
        {"Name", each if try Record.Field(_, "MultiRole") <> null otherwise false then _ & "(" & Record.Field(_, "MultiRole") & ")" else _}
    ),
    #"Renamed Columns" = Table.RenameColumns(#"Transformed Columns",{{"Name", "NameX"}}),
    #"NEW NAME" = Table.AddColumn(#"Renamed Columns", "Name", each if [MultiRole] <> null then [NameX] & " (" & [MultiRole] & ")" else [NameX]),
    #"Renamed Columns1" = Table.RenameColumns(#"NEW NAME",{{"Role", "RoleX"}}),
    #"Added Conditional Column" = Table.AddColumn(#"Renamed Columns1", "Role", each if [MultiRole] <> null then [MultiRole] else [RoleX]),
    #"Removed Columns" = Table.RemoveColumns(#"Added Conditional Column",{"AINC4HrsAvilPref", "MultiRole", "RoleSplit", "RoleX", "NameX"})
in
    #"Removed Columns";

shared RoleAvailabilityCapped = let
    Source = RoleResDayAvailabilityCapped,
    #"Grouped Rows" = Table.Group(Source, {"Role"}, {{"AvailabilityCapped", each List.Sum([AvailabilityCapped]), type number}})
in
    #"Grouped Rows";

shared #"Availability-StaffList" = let
    Source = RoleResDayAvailabilityCapped,
    #"Grouped Rows" = Table.Group(Source, {"Name", "Role"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Name", Order.Ascending}})
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

shared #"FilePath - 2Calculations" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

shared #"FilePath - 1Input" = let
    Source = #"UnitL1PathTABLE",
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows","2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{0}[Value]
in
    Value;