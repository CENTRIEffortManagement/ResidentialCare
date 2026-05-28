// Power Query from: StaffListMaster.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\StaffListMaster.xlsx
// Extracted: 2026-05-18T06:14:39.153Z

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

shared #"FilePath-2Calculations" = let
    Source = UnitL1PathTABLE,
    Value = Source{1}[Value]
in
    Value;

shared #"IMPORT AvailabilityDayShiftMATRIXRaw" = let

    Source = Excel.Workbook(File.Contents(#"FilePath-2Calculations" & "\Capacity-ShiftAvailability.xlsx"), null, true),
    #"Availability-StaffList_Sheet" = Source{[Item="Availability-StaffList",Kind="Sheet"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(#"Availability-StaffList_Sheet",{{"Column1", type text}, {"Column2", type text}}),
    #"Promoted Headers" = Table.PromoteHeaders(#"Changed Type", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"Name", type text}, {"Role", type text}})
in
    #"Changed Type1";

shared #"Masterlist Join" = let
    Source = AvailableStaffList,
    #"Appended Query" = Table.Combine({Source, AllocationTable_StaffList}),
    #"Removed Columns" = Table.RemoveColumns(#"Appended Query",{"Source"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns", {"Name"})
in
    #"Removed Duplicates";

shared Masterlist = let
    Source = #"Masterlist Join",
    #"Added RESOURCE INDEX" = Table.AddIndexColumn(Source, "Resource", 1, 1, Int64.Type),
    #"Merged Queries" = Table.NestedJoin(#"Added RESOURCE INDEX", {"Name", "Role"}, #"Misaligned Join", {"Misaligned-Name", "Misaligned-Role"}, "Misaligned Join", JoinKind.FullOuter),
    #"Expanded Misaligned Join" = Table.ExpandTableColumn(#"Merged Queries", "Misaligned Join", {"Misalignment"}, {"Misalignment"})
in
    #"Expanded Misaligned Join";

shared AllocationTable_StaffList = let
     Source1 = #"IMPORT Table_AllocatedStaffList",
    #"Sorted Rows" = Table.Sort(Source1,{{"Name", Order.Ascending}}),
    #"Added Custom" = Table.AddColumn(#"Sorted Rows", "Source", each "Allocation")
in
    #"Added Custom";

shared AvailableStaffListInitial = let
    Source = #"IMPORT AvailabilityDayShiftMATRIXRaw",
    #"Grouped Rows" = Table.Group(Source, {"Role", "Name"}, {{"AvailableListInitial", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"AvailableListInitial"})
in
    #"Removed Columns";

shared AvailableStaffList = let
     Source1 = AvailableStaffListInitial,
    #"Added Custom" = Table.AddColumn(Source1, "Source", each "Availability"),
    #"Duplicated Column" = Table.DuplicateColumn(#"Added Custom", "Name", "Name.A"),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Duplicated Column", "Name", Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv), {"Name.1", "Name.2", "Name.3", "Name.4", "Name.5", "Name.6"}),
    #"Removed Columns" = Table.RemoveColumns(#"Split Column by Delimiter",{"Name.4", "Name.5", "Name.6"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns",null,"",Replacer.ReplaceValue,{"Name.2", "Name.3"}),
    #"Replaced NANDEEENI" = Table.ReplaceValue(#"Replaced Value","Muni","",Replacer.ReplaceText,{"Name.2"}),
    #"Merged Columns" = Table.CombineColumns(#"Replaced NANDEEENI",{"Name.2", "Name.3"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Merged"),
    #"Trimmed Text" = Table.TransformColumns(#"Merged Columns",{{"Merged", Text.Trim, type text}}),
    #"Merged Columns1" = Table.CombineColumns(#"Trimmed Text",{"Name.1", "Merged"},Combiner.CombineTextByDelimiter(", ", QuoteStyle.None),"Name"),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Merged Columns1", "Name", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Name.1", "Name.2"}),
    #"Renamed Columns" = Table.RenameColumns(#"Split Column by Delimiter1",{{"Name.A", "Name"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns",{{"Name", Order.Ascending}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Sorted Rows",{"Name.1", "Name.2"})
in
    #"Removed Columns1";

shared #"IMPORT Table_AllocatedStaffList" = let
    Source = Excel.Workbook(File.Contents(#"FilePath-1Input"& "\1-AllocationExtracted.xlsx"), null, true),
    Table_AllocatedStaffList_Table = Source{[Item="Table_AllocatedStaffList",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table_AllocatedStaffList_Table,{{"Name", type text}, {"Role", type text}})
in
    #"Changed Type";

shared #"Misaligned Join" = let
    Source = Table.NestedJoin(AllocationTable_StaffList, {"Name", "Role"}, AvailableStaffList, {"Name", "Role"}, "Availability-StaffList", JoinKind.FullOuter),
    #"Expanded Availability-StaffList" = Table.ExpandTableColumn(Source, "Availability-StaffList", {"Name", "Role"}, {"Availability-StaffList.Name", "Availability-StaffList.Role"}),
    #"Sorted Rows" = Table.Sort(#"Expanded Availability-StaffList",{{"Name", Order.Ascending}}),
    #"Added Conditional Column" = Table.AddColumn(#"Sorted Rows", "Custom", each if [Name] <> [#"Availability-StaffList.Name"] then "Misaligned" else null),
    #"Sorted Rows1" = Table.Sort(#"Added Conditional Column",{{"Availability-StaffList.Name", Order.Ascending}}),
    #"Filtered Rows1" = Table.SelectRows(#"Sorted Rows1", each ([Custom] = "Misaligned")),
    #"Inserted Merged Column" = Table.AddColumn(#"Filtered Rows1", "Misaligned-Name", each Text.Combine({[Name], [#"Availability-StaffList.Name"]}, ""), type text),
    #"Inserted Merged Column1" = Table.AddColumn(#"Inserted Merged Column", "Misaligned-Role", each Text.Combine({[Role], [#"Availability-StaffList.Role"]}, ""), type text),
    #"Added Conditional Column1" = Table.AddColumn(#"Inserted Merged Column1", "Misalignment", each if [Name] = null then "Available Only" else if [#"Availability-StaffList.Name"] = null then "Allocated Only" else "ERROR")
in
    #"Added Conditional Column1";

shared #"Misaligned Allocation Availability Names" = let
    Source = #"Misaligned Join",
    #"Merged Queries" = Table.NestedJoin(Source, {"Misaligned-Name", "Misaligned-Role"}, Masterlist, {"Name", "Role"}, "Masterlist", JoinKind.LeftOuter),
    #"Expanded Masterlist" = Table.ExpandTableColumn(#"Merged Queries", "Masterlist", {"Resource"}, {"Resource"}),
    #"Sorted Rows2" = Table.Sort(#"Expanded Masterlist",{{"Misaligned-Role", Order.Ascending}, {"Resource", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows2",{"Custom", "Misaligned-Name", "Misaligned-Role", "Misalignment", "Resource"})
in
    #"Removed Other Columns";

shared #"FilePath-1Input" = let
    Source = UnitL1PathTABLE,
    #"Replaced Value" = Table.ReplaceValue(Source,"2. Calculations","1. Input",Replacer.ReplaceText,{"Value"}),
    Value = #"Replaced Value"{1}[Value]
in
    Value;