// Power Query from: Effort-All.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\E-O-I\Effort-All.xlsx
// Extracted: 2026-09-13T06:13:33.983Z

section Section1;

// Query: UnitL1PathTABLE
// Purpose: Resolve this client-level workbook through the standard CentriSyncPaths mapping.
// Inputs: Canonical FilePathUrl input with one FilePath value; a single Column1 named cell is also supported.
// Output: Existing Variable Name / Value interface, plus client and units roots for imports.
shared UnitL1PathTABLE = let
    NormalizePath = (value as nullable text) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(value), "/", "\"), "\"),
    PathInput = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
    // Retain the single named-cell shape supported by the previous resolver, but reject unrelated columns.
    PathColumn = if Table.HasColumns(PathInput, {"FilePath"}) then Table.SelectColumns(PathInput, {"FilePath"})
        else if Table.ColumnNames(PathInput) = {"Column1"} then Table.RenameColumns(PathInput, {{"Column1", "FilePath"}})
        else error "FilePathUrl must contain a FilePath column or a single Column1 named cell.",
    SinglePath = if Table.RowCount(PathColumn) = 1 then PathColumn{0}[FilePath]
        else error "FilePathUrl must contain exactly one data row.",
    RawFilePath = if not Value.Is(SinglePath, type text) then error "FilePathUrl must contain a saved workbook path as text."
        else if Text.Trim(SinglePath) = "" then error "FilePathUrl is blank. Save this workbook and recalculate its filename formula."
        else NormalizePath(SinglePath),
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet; strip the sheet suffix.
    WorkbookPath = if Text.Contains(RawFilePath, "[") then
        Text.BeforeDelimiter(RawFilePath, "[") & Text.BetweenDelimiters(RawFilePath, "[", "]")
        else RawFilePath,
    FileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    ValidatedPath = if Comparer.OrdinalIgnoreCase(FileName, "Effort-All.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Save Effort-All.xlsx and recalculate its filename formula.",
    // Fixed public-machine configuration location; no user-specific source paths.
    MappingWorkbook = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    MappingTable = MappingWorkbook{[Item="CentriSyncPaths", Kind="Table"]}[Data],
    MappingTypes = Table.TransformColumnTypes(
        Table.SelectColumns(MappingTable, {"SharepointRootUrl", "SyncedFolderRootPath"}),
        {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizedMappings = Table.TransformColumns(MappingTypes,
        {{"SharepointRootUrl", NormalizePath, type nullable text}, {"SyncedFolderRootPath", NormalizePath, type nullable text}}),
    SiteCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    DocumentCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot",
        each if [SharepointRootUrl] = null or [SharepointRootUrl] = "" then null else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    Candidates = Table.AddColumn(Table.Combine({SiteCandidates, DocumentCandidates, LocalCandidates}),
        "RootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    // Match complete path segments, so similarly named sites or local folders cannot capture this workbook.
    Matches = Table.SelectRows(Candidates, each [RootLength] > 0
        and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and (Comparer.OrdinalIgnoreCase(ValidatedPath, [MatchRoot]) = 0
            or Text.StartsWith(ValidatedPath, [MatchRoot] & "\", Comparer.OrdinalIgnoreCase))),
    SortedMatches = Table.Sort(Matches, {{"RootLength", Order.Descending}}),
    LongestMatches = if Table.IsEmpty(SortedMatches) then error "FilePathUrl did not match any CentriSyncPaths root: " & ValidatedPath
        else Table.SelectRows(SortedMatches, each [RootLength] = SortedMatches{0}[RootLength]),
    // Conflicting mappings at equal specificity must fail instead of depending on spreadsheet row order.
    Destinations = List.Distinct(LongestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    BestMatch = if List.Count(Destinations) = 1 then LongestMatches{0}
        else error "CentriSyncPaths contains conflicting mappings for this workbook.",
    MappedRoot = BestMatch[SyncedFolderRootPath],
    LocalRoot = if (Text.Length(MappedRoot) >= 3 and Text.Range(MappedRoot, 1, 2) = ":\")
        or Text.StartsWith(MappedRoot, "\\") then MappedRoot
        else error "CentriSyncPaths SyncedFolderRootPath must be an absolute local or UNC path.",
    LocalFullPath = LocalRoot & Text.Range(ValidatedPath, BestMatch[RootLength]),
    WorkbookFolder = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    // E-O-I is under the client, outside UNITS. Derive the client root from its exact folder suffix.
    CalculationSuffix = "\2. Calculations\E-O-I",
    ClientRoot = if Text.EndsWith(WorkbookFolder, CalculationSuffix, Comparer.OrdinalIgnoreCase) then
        Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(CalculationSuffix))
        else error "Effort-All.xlsx must be saved under the client's 2. Calculations\E-O-I folder.",
    Segments = List.Select(Text.Split(WorkbookFolder, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(WorkbookFolder, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    Date = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    // Buffer the small path table because multiple scalar paths consume the same mapping result.
    PathTable = #table(type table [#"Variable Name" = text, Value = nullable text], {
        {"UserName", UserName}, {"Root Path", WorkbookFolder}, {"FilePathUrl", RawFilePath},
        {"Client", Client}, {"Date", Date}, {"Unit", null}, {"FileName", FileName},
        {"Client Path", ClientRoot}, {"Units Path", ClientRoot & "\UNITS"}
    })
in
    Table.Buffer(PathTable);

// Query: UnitsPath
// Purpose: Expose the client's UNITS root for configured facility imports.
shared UnitsPath = UnitL1PathTABLE{[#"Variable Name"="Units Path"]}[Value];

// Query: ResolveFacilityPath
// Purpose: Resolve a configured unit folder beneath the client's UNITS root.
// Inputs: Existing Table3 Facility1/Facility2 values, such as Unit1 or UNITS\Unit1.
shared ResolveFacilityPath = (facility as any) as text => let
    Normalized = if Value.Is(facility, type text) then Text.Trim(Text.Replace(Text.Trim(facility), "/", "\"), "\")
        else error "Table3 facility paths must be nonblank text identifying a unit folder.",
    UnitName = if Text.StartsWith(Normalized, "UNITS\", Comparer.OrdinalIgnoreCase) then Text.Range(Normalized, 6) else Normalized,
    // A single folder name prevents absolute paths and traversal outside the resolved client.
    ValidatedUnit = if UnitName = "" or List.Contains({".", ".."}, UnitName)
        or Text.Contains(UnitName, "\") or Text.Contains(UnitName, ":") then
            error "Table3 Facility1/Facility2 must identify a unit folder, for example Unit1 or UNITS\Unit1."
        else UnitName
in
    UnitsPath & "\" & ValidatedUnit;

shared #"INPUT Facilities" = let
    Source = Excel.CurrentWorkbook(){[Name="Table3"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Facility1", type text}, {"Facility2", type text}})
in
    #"Changed Type";

shared Facility1 = let
    Source = #"INPUT Facilities",
    Facility1 = Source{0}[Facility1]
in
    Facility1;

shared Facility2 = let
    Source = #"INPUT Facilities",
    Facility1 = Source{0}[Facility2]
in
    Facility1;

// Query: Facility1 Effort Workbook
// Purpose: Share the first configured unit's Effort workbook across its table imports.
// Notes: Buffer the file binary once per evaluation to avoid repeated reads within the import.
shared #"IMPORTFacility1 Effort Workbook" = Excel.Workbook(Binary.Buffer(File.Contents(
    ResolveFacilityPath(Facility1) & "\2. Calculations\Effort.xlsx")), null, true);

// Query: Facility1-EffortAllMatrixAG1_1D
// Purpose: Import the first configured unit's effort matrix from the shared workbook source.
shared #"Facility1-EffortAllMatrixAG1_1D" = let
    Source = #"IMPORTFacility1 Effort Workbook",
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_Table,{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text},  {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"CapacityMaxHC", Int64.Type}, {"Allocation", type number}, {"DATESHIFT", type text}})
in
    #"Changed Type";

// Query: Facility1 - Availabilities
// Purpose: Import the first configured unit's availability rows from the shared workbook source.
shared #"Facility1 - Availabilities" = let
    Source = #"IMPORTFacility1 Effort Workbook",
    RoleShiftAvailabilities_Table = Source{[Item="RoleShiftAvailabilities",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftAvailabilities_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"Capacity", type text}, {"Facility", type text}, {"Date", type date}, {"Shift", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}, {"Resource", Order.Ascending}})
in
    #"Sorted Rows";

// Query: Facility1 - ResAllocation
// Purpose: Import the first configured unit's resource allocation from the shared workbook source.
shared #"Facility1 - ResAllocation" = let
    Source = #"IMPORTFacility1 Effort Workbook",
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type1",null,43,Replacer.ReplaceValue,{"Period"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Replaced Value",{{"Period", Int64.Type}})
in
    #"Changed Type";

// Query: Facility2 Effort Workbook
// Purpose: Share the second configured unit's Effort workbook across its table imports.
// Notes: Buffer the file binary once per evaluation; reuse the first import if both paths are equal.
shared #"IMPORT Facility2 Effort Workbook" = if Comparer.OrdinalIgnoreCase(ResolveFacilityPath(Facility1), ResolveFacilityPath(Facility2)) = 0 then
    #"IMPORTFacility1 Effort Workbook"
    else Excel.Workbook(Binary.Buffer(File.Contents(ResolveFacilityPath(Facility2) & "\2. Calculations\Effort.xlsx")), null, true);

// Query: Facility2-EffortAllMatrixAG1_1D
// Purpose: Import and date-align the second configured unit's effort matrix.
shared #"Facility2-EffortAllMatrixAG1_1D#" = let
    Source = #"IMPORT Facility2 Effort Workbook",
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Renamed Columns" = Table.RenameColumns(EffortAllMatrixAG1_1D_Table,{{"Date", "Datetemp"}}),
    #"Added ALIGNWITHBERR" = Table.AddColumn(#"Renamed Columns", "Custom", each Date.AddDays([Datetemp],#"INPUT DateAlignment")),
    #"Renamed Columns1" = Table.RenameColumns(#"Added ALIGNWITHBERR",{{"Custom", "Date"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns1",{"Facility", "Role", "Date", "Datetemp", "Shift", "Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation", "DATESHIFT"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Datetemp"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "ShiftDurations.ShiftDuration", each 1),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Date", type date}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type","1","2",Replacer.ReplaceText,{"Facility"})
in
    #"Replaced Value";

// Query: Facility2 - Availabilities ##
// Purpose: Import the second configured unit's availability rows from the shared workbook source.
shared #"Facility2 - Availabilities ##" = let
    Source = #"IMPORT Facility2 Effort Workbook",
    RoleShiftAvailabilities_Table = Source{[Item="RoleShiftAvailabilities",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftAvailabilities_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"Capacity", type text},  {"Date", type date}, {"Shift", type text}})
in
    #"Changed Type1";

// Query: Facility2 - ResAllocation
// Purpose: Import the second configured unit's resource allocation from the shared workbook source.
shared #"Facility2 - ResAllocation" = let
    Source = #"IMPORT Facility2 Effort Workbook",
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",null,43,Replacer.ReplaceValue,{"Period"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"Period", Int64.Type}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type1","1","2",Replacer.ReplaceText,{"Facility"})
in
    #"Replaced Value1";

// Query: ResShiftAllocation1
// Purpose: Import and sort the second configured unit's resource allocation.
shared #"Facility2 ResShiftAllocation1" = let
    Source = #"IMPORT Facility2 Effort Workbook",
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Role", Order.Ascending}, {"Resource", Order.Ascending}, {"TimeDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    #"Sorted Rows";

shared #"INPUT DateAlignment" = let
    Source = Excel.CurrentWorkbook(){[Name="DateAlignment"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"DateAlignment", Int64.Type}}),
    DateAlignment1 = #"Changed Type"{0}[DateAlignment]
in
    DateAlignment1;

// Query: Folder
// Purpose: Preserve the scalar Folder interface as the resolved client root, without a trailing separator.
shared Folder = UnitL1PathTABLE{[#"Variable Name"="Client Path"]}[Value];

// Query: Table_RevisedAllocated
// Purpose: Import allocation changes relative to the resolved client root.
shared #"IMPORT Table_RevisedAllocated" = let

    
    Source = Excel.Workbook(File.Contents(Folder & "\2. Calculations\Change\AllocationChange.xlsx"), null, true),
    Table_RevisedAllocated_Table = Source{[Item="Table_RevisedAllocated",Kind="Table"]}[Data],
    #"Added Custom" = Table.AddColumn(Table_RevisedAllocated_Table, "Version", each "ChangeAllcoation")
in
    #"Added Custom";

shared EffectiveAvailability = 7.6/8 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared ResShiftAllocation = let
    Source1 = #"Facility2 - ResAllocation",
    #"Appended Query" = Table.Combine({Source1, #"Facility1 - ResAllocation"}),
    //CombinedTables = Table.Combine({Source1, #"Facility1 - ResAllocation"}),
    #"Renamed Columns" = Table.RenameColumns(#"Appended Query",{{"ShiftPeriod", "Shift"}, {"TimeDate", "Date"}, {"ResShiftFTE", "Effort"}, {"Type", "EffortType"}}),
    #"Filtered RES NULL" = Table.SelectRows(#"Renamed Columns", each ([Resource] <> null))
in
    #"Filtered RES NULL";

shared #"EffortAllMatrixAG1_1D-base !!" = let
   Source = #"Facility2-EffortAllMatrixAG1_1D#",
    #"Appended Query" = Table.Combine({Source, #"Facility1-EffortAllMatrixAG1_1D"}),
    #"Multiplied Column" = Table.TransformColumns(#"Appended Query", {{"Capacity", each _ * EffectiveAvailability, type number}}),
    #"Multiplied Column1" = Table.TransformColumns(#"Multiplied Column", {{"CapacityX", each _ * EffectiveAvailability, type number}}),
    //#"Appended Query" = Table.Combine({Source, #"Facility1-EffortAllMatrixAG1_1D"}),
    #"Renamed Columns" = Table.RenameColumns(#"Multiplied Column1",{{"Date", "DateX"}}),  // Add back Facility 2
    #"Added DATE ALIGNMENT" = Table.AddColumn(#"Renamed Columns", "Date", each if [Facility] = "ASHB" then  Date.AddDays([DateX],-#"INPUT DateAlignment")
else [DateX]),
    #"Removed Columns" = Table.RemoveColumns(#"Added DATE ALIGNMENT",{"DateX"})
in
    #"Removed Columns";

shared #"EffortAllMatrixAG1_1D-Revised" = let
    Source = #table({}, {})
    
    // Revised allocation remains disabled; preserve the existing empty output.

    in
    Source;

shared EffortAllMatrixAG1_1D = let
    Source = #"EffortAllMatrixAG1_1D-base !!",
    // = Table.Combine({#"EffortAllMatrixAG1_1D-base !!", #"EffortAllMatrixAG1_1D-Revised"}),
    #"Sorted Rows1" = Table.Sort(Source,{{"DATESHIFT", Order.Ascending}}),
    #"Sorted Rows" = Table.Sort(#"Sorted Rows1",{{"DATESHIFT", Order.Ascending}}),
    #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each ([Role] <> null)),
    #"Sorted Rows2" = Table.Sort(#"Filtered Rows",{{"Facility", Order.Ascending}, {"Date", Order.Ascending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows2",{{"Date", type date}}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Changed Type", "Date", "Date - Copy"),
    #"Renamed Columns" = Table.RenameColumns(#"Duplicated Column",{{"Date - Copy", "DateX"}})
in
    #"Renamed Columns";

shared #"Availabilities Appended" = let
    Source = #"Facility2 - Availabilities ##",
    #"Appended FAC1" = Table.Combine({Source, #"Facility1 - Availabilities"}),
    //Table.Combine({#"Facility1 - Availabilities", #"Facility2 - Availabilities"}),    //return Facility 1
    #"Renamed Columns" = Table.RenameColumns(#"Appended FAC1",{{"Capacity", "EffortType"}, {"Availability", "Effort"}}),
    #"Multiplied Column" = Table.TransformColumns(#"Renamed Columns", {{"Effort", each _ * EffectiveAvailability, type number}}),
    #"Sorted Rows" = Table.Sort(#"Multiplied Column",{{"Facility", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}, {"Resource", Order.Ascending}, {"EffortType", Order.Ascending}}),
    #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each ([Resource] = 108))
in
    #"Filtered Rows";

shared #"Availability Effort" = let
    Source = #"EffortAllMatrixAG1_1D-base !!",
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Facility", "Role", "Shift", "Demand", "Date"}),
    #"Added DEMAND" = Table.AddColumn(#"Removed Other Columns", "EffortType", each "Demand"),
    #"Renamed Columns" = Table.RenameColumns(#"Added DEMAND",{{"Demand", "Effort"}}),
    #"Appended AVAILABILITIES" = Table.Combine({#"Renamed Columns", #"Availabilities Appended"}),
    #"Removed Columns" = Table.RemoveColumns(#"Appended AVAILABILITIES",{"ResAvailability", "ResMaxAvail"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Columns",{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text}, {"EffortType", type text}, {"Effort", type number}, {"Resource", Int64.Type}}),
    #"Appended ALLOCATION" = Table.Combine({#"Changed Type", ResShiftAllocation}),
    #"Replaced Value" = Table.ReplaceValue(#"Appended ALLOCATION",0,null,Replacer.ReplaceValue,{"Effort"}),
    #"Sorted Rows" = Table.Sort(#"Replaced Value",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Resource", Order.Ascending}})
in
    #"Sorted Rows";

shared ResRoleAvailabilityDevelopedMATRIX = let
    Source = #"Availability Effort",
    #"Sorted Rows2" = Table.Sort(Source,{{"Date", Order.Ascending}, {"EffortType", Order.Ascending}}),
    #"Sorted Rows" = Table.Sort(#"Sorted Rows2",{{"Effort", Order.Descending}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Sorted Rows",{{"Effort", type number}}),
    #"Sorted Rows1" = Table.Sort(#"Changed Type1",{{"Resource", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}, {"EffortType", Order.Ascending}}),
    #"Pivoted Column" = Table.Pivot(#"Sorted Rows1", List.Distinct(#"Sorted Rows1"[EffortType]), "EffortType", "Effort")
in
    #"Pivoted Column";

shared #"Capacity SUM" = let
    Source = EffortAllMatrixAG1_1D,
    //#"Filtered Rows" = Table.SelectRows(Source, each ([Version] = "Base")),
    #"Grouped Rows" = Table.Group(#"Source", {"Facility", "Role"}, {{"Capacity", each List.Sum([Capacity]), type number}}),
    #"Replaced Value" = Table.ReplaceValue(#"Grouped Rows",null,"?",Replacer.ReplaceValue,{"Role"}),
    Custom1 = List.Sum(#"Replaced Value"[Capacity]),
    #"Converted to Table" = #table(1, {{Custom1}}),
    #"Renamed Columns1" = Table.RenameColumns(#"Converted to Table",{{"Column1", "TOTAL"}}),
    TOTAL = #"Renamed Columns1"{0}[TOTAL]
in
    TOTAL;

shared RoleAvailabilityDevelopedMATRIXDELTA = let
    Source = ResRoleAvailabilityDevelopedMATRIX,
    #"Grouped Rows" = Table.Group(Source, {"Facility", "Role", "Shift", "Date"}, {{"Demand", each List.Sum([Demand]), type nullable number}, {"C###", each List.Sum([#"C###"]), type nullable number}, {"Allocation", each List.Sum([Allocation]), type nullable number}, {"Original Availability", each List.Sum([OriginalAvailability]), type nullable number}}),
    #"Inserted Subtraction" = Table.AddColumn(#"Grouped Rows", "C###-D", each [#"C###"] - [Demand], type number),
    #"Inserted Addition" = Table.AddColumn(#"Inserted Subtraction", "A-D", each [Allocation] - [Demand], type number),
    #"Inserted Subtraction1" = Table.AddColumn(#"Inserted Addition", "A-C###", each [#"C###"] - [Allocation], type number),
    #"Inserted Subtraction2" = Table.AddColumn(#"Inserted Subtraction1", "CO-D", each [Original Availability] - [Demand], type number)
in
    #"Inserted Subtraction2";

shared MaxAvailabilities = let
    Source = #"Availabilities Appended",
    #"Filtered Rows" = Table.SelectRows(Source, each ([EffortType] = "C###")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Role", "Resource", "Effort", "Facility", "ResAvailability", "ResMaxAvail"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"Effort", "C###"}})
in
    #"Renamed Columns";