// Power Query from: Capacity.xlsx
// Pathname: c:\Users\Cliff's Computer\Centri\3. Product - Documents\mcode Dev\ResidentialCare\CLIENT\DATExx\UNITS\Unit1\2. Calculations\Capacity.xlsx
// Extracted: 2026-05-18T06:14:12.971Z

section Section1;

shared #"IMPORT C#TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C#"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"RosteredPeriodStatus"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"AvailabilityCapped", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT C##TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT C###TABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&Role2&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE2" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role2"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

shared UPLOADRole = let
    Source = Excel.CurrentWorkbook(){[Name="Table6"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"RoleID", type text}, {"RoleName", type text}})
in
    #"Changed Type";

shared Role1 = let
    Source = UPLOADRole,
    RoleName = Source{0}[RoleName]
in
    RoleName;

shared Role2 = let
    Source = UPLOADRole,
    RoleName = Source{1}[RoleName]
in
    RoleName;

shared Role3 = let
    Source = UPLOADRole,
    RoleName = Source{2}[RoleName]
in
    RoleName;

shared Folder = let
    Source = Unit1Path
in
    Source;

shared Facility = let
    Source = UnitL1PathTABLE,
    Value = Source{5}[Value]
in
    Value;

shared #"IMPORT Role1C#TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

shared #"IMPORT Role1C##TABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT Role1C###TABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder & "\" & #"Role1"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Role] <> null )),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE1" = let
    Source = Excel.Workbook(File.Contents(Folder &"\" & #"Role1"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

shared Capacity = let
    Source = #"IMPORT Role1C###TABLE1",
    #"Appended Query" = Table.Combine({Source, #"IMPORT C###TABLE2", #"IMPORT Role3C###TABLE"}),
    #"Filtered Rows" = Table.SelectRows(#"Appended Query", each ([Resource] = 94)),
    #"Merged Queries" = Table.NestedJoin(#"Filtered Rows", {"Period", "Role"}, PermutationDimensions, {"Period", "RolesList"}, "PermutationDimensions", JoinKind.LeftOuter),
    #"Expanded PermutationDimensions" = Table.ExpandTableColumn(#"Merged Queries", "PermutationDimensions", {"Date"}, {"Date"}),
    #"Sorted Rows" = Table.Sort(#"Expanded PermutationDimensions",{{"Role", Order.Ascending}, {"Resource", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows"
   ;

shared AvailabilityDeveloped = let
    Source = Table.Combine({#"IMPORT OriginalResPeriodAvailabilityTABLE1", #"IMPORT Role1C#TABLE", #"IMPORT Role1C##TABLE1", #"IMPORT Role1C###TABLE1", #"IMPORT C#TABLE2", #"IMPORT C##TABLE2", #"IMPORT C###TABLE2", #"IMPORT Role3C#TABLE", #"IMPORT Role3C##TABLE", #"IMPORT Role3C###TABLE", #"IMPORT OriginalResPeriodAvailabilityTABLE2", #"IMPORT OriginalResPeriodAvailabilityTABLE3"}),
    #"Added Custom" = Table.AddColumn(Source, "Facility", each Facility),
    #"Replaced Value" = Table.ReplaceValue(#"Added Custom",0,null,Replacer.ReplaceValue,{"Availability"})
in
    #"Replaced Value";

shared AvailabilityDevelopedMATRIX = let
    Source = AvailabilityDeveloped,
    #"Removed Columns" = Table.RemoveColumns(Source,{"ResAvailability", "ResMaxAvail"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[AvailabilityType]), "AvailabilityType", "Availability", List.Sum)
in
    #"Pivoted Column";

shared PermutationDimensions = let
    Source = Excel.Workbook(File.Contents(Folder&"\Settings Data.xlsx"), null, true),
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Date", Order.Ascending}, {"Period", Order.Ascending}})
in
    #"Sorted Rows";

shared #"IMPORT Role3C#TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(A.2)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityCapped_C__TABLE_Table = Source{[Item="ResPeriodAvailabilityCapped_C__TABLE",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResPeriodAvailabilityCapped_C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"RosteredPeriodStatus", type any}, {"AvailabilityCapped", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type1", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C#"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"AvailabilityCapped", "Availability"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"RosteredPeriodStatus"})
in
    #"Removed Columns";

shared #"IMPORT Role3C##TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C__TABLE_Table = Source{[Item="C__TABLE",Kind="Table"]}[Data],
    #"Changed Type2" = Table.TransformColumnTypes(C__TABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C##", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type2", each ([Role] <> null)),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C##"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C##", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT Role3C###TABLE" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(B)-shifts.xlsx"), null, true),
    C___TABLE_B_Table = Source{[Item="C___TABLE_B",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(C___TABLE_B_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"C###", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Role] <> null )),
    #"Added Custom" = Table.AddColumn(#"Filtered Rows", "AvailabilityType", each "C###"),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"C###", "Availability"}})
in
    #"Renamed Columns";

shared #"IMPORT OriginalResPeriodAvailabilityTABLE3" = let
    Source = Excel.Workbook(File.Contents(Folder&"\"&#"Role3"&"\CapacityDistrib(A.1)-shifts.xlsx"), null, true),
    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "AvailabilityType", each "OriginalAvailability")
in
    #"Added Custom";

shared UnitL1PathTABLE = // Version 25.00    250107 
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
    )
in
    TABLE;

shared Unit1Path = let
    Source = UnitL1PathTABLE,
    Value = Source{1}[Value]
in
    Value;