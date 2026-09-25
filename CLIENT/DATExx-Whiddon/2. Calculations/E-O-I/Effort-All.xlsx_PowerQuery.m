// Power Query from: Effort-All.xlsx
// Pathname: c:\Users\Alex\CentriNOTSYNC\ResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\E-O-I\Effort-All.xlsx
// Extracted: 2026-09-20T21:19:08.859Z

section Section1;

// Query: IMPORT CentriSyncPaths
// Purpose: Import the public-machine ResidentialCare path mappings once for the authoritative resolver.
shared #"IMPORT CentriSyncPaths" = let
    WorkbookNavigation = Excel.Workbook(Binary.Buffer(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx")), null, true),
    Matches = Table.SelectRows(WorkbookNavigation, each [Item] = "CentriSyncPaths" and [Kind] = "Table"),
    MappingTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Effort-All path resolver", "Expected exactly one CentriSyncPaths table in the public mapping workbook.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"SharepointRootUrl", "SyncedFolderRootPath"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(MappingTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(MappingTable, RequiredColumns)
        else error Error.Record("Effort-All path resolver", "CentriSyncPaths is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected, {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}})
in
    Table.Buffer(Typed);

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
    // The fixed public-machine configuration is opened only by IMPORT CentriSyncPaths.
    MappingTypes = #"IMPORT CentriSyncPaths",
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
// Inputs: EffortAllSettings Facility1/Facility2 values, such as Unit1 or UNITS\Unit1.
shared ResolveFacilityPath = (facility as any) as text => let
    Normalized = if Value.Is(facility, type text) then Text.Trim(Text.Replace(Text.Trim(facility), "/", "\"), "\")
        else error "EffortAllSettings facility paths must be nonblank text identifying a unit folder.",
    UnitName = if Text.StartsWith(Normalized, "UNITS\", Comparer.OrdinalIgnoreCase) then Text.Range(Normalized, 6) else Normalized,
    // A single folder name prevents absolute paths and traversal outside the resolved client.
    ValidatedUnit = if UnitName = "" or List.Contains({".", ".."}, UnitName)
        or Text.Contains(UnitName, "\") or Text.Contains(UnitName, ":") then
            error "EffortAllSettings Facility1/Facility2 must identify a unit folder, for example Unit1 or UNITS\Unit1."
        else UnitName
in
    UnitsPath & "\" & ValidatedUnit;

shared #"Facility Analysis" = let
    Source = #"LINK Unit Analysis",
    #"Filtered UNITS" = Table.SelectRows(Source, each ([Effort Management Analysis] = true)),
    #"Added Index" = Table.AddIndexColumn(#"Filtered UNITS", "Index", 1, 1, Int64.Type),
    #"Added Prefix" = Table.TransformColumns(#"Added Index", {{"Index", each "Facility" & Text.From(_, "en-AU"), type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Added Prefix",{"Effort Management Analysis"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[Index]), "Index", "Title")
in
    #"Pivoted Column";

// Query: IMPORT AllocationChange
// Purpose: Import and buffer AllocationChange.xlsx once relative to the resolved client root.
shared #"IMPORT AllocationChange" = Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
    Folder & "\2. Calculations\Change\AllocationChange.xlsx")), null, true));

// Query: EXTRACT Table_RevisedAllocated
// Purpose: Extract the allocation-change table from the shared AllocationChange import.
shared #"EXTRACT Table_RevisedAllocated" = let
    Source = #"IMPORT AllocationChange",
    Table_RevisedAllocated_Table = Source{[Item="Table_RevisedAllocated",Kind="Table"]}[Data],
    #"Added Custom" = Table.AddColumn(Table_RevisedAllocated_Table, "Version", each "ChangeAllcoation")
in
    #"Added Custom";

// Query: IMPORT Facility1 Effort
// Purpose: Share the first configured unit's Effort workbook across its table imports.
// Notes: Buffer the file binary once per evaluation to avoid repeated reads within the import.
shared #"IMPORT Facility1 Effort" = Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
    ResolveFacilityPath(Units{0}[Facility1]) & "\2. Calculations\Effort.xlsx")), null, true));

// Query: EXTRACT Facility1 EffortAllMatrixAG1_1D
// Purpose: Extract the first configured unit's effort matrix from the shared Effort import.
shared #"EXTRACT Facility1 EffortAllMatrixAG1_1D" = let
    Source = #"IMPORT Facility1 Effort",
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_Table,{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text},  {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"CapacityMaxHC", Int64.Type}, {"Allocation", type number}, {"DATESHIFT", type text}})
in
    #"Changed Type";

// Query: EXTRACT Facility1 RoleShiftAvailabilities
// Purpose: Extract the first configured unit's availability rows from the shared Effort import.
shared #"EXTRACT Facility1 RoleShiftAvailabilities" = let
    Source = #"IMPORT Facility1 Effort",
    RoleShiftAvailabilities_Table = Source{[Item="RoleShiftAvailabilities",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftAvailabilities_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"Capacity", type text}, {"Facility", type text}, {"Date", type date}, {"Shift", type text}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"Facility", Order.Ascending}, {"Role", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}, {"Resource", Order.Ascending}})
in
    #"Sorted Rows";

// Query: EXTRACT Facility1 ResShiftAllocation
// Purpose: Extract the first configured unit's resource allocation from the shared Effort import.
shared #"EXTRACT Facility1 ResShiftAllocation" = let
    Source = #"IMPORT Facility1 Effort",
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type1",null,43,Replacer.ReplaceValue,{"Period"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Replaced Value",{{"Period", Int64.Type}})
in
    #"Changed Type";

// Query: EXTRACT Facility1 ResourceContractOverage
// Purpose: Extract the Unit1/AIN employee-level contract overage table from the first configured Effort workbook.
// Output: One row per employee, preferred Resource and roster; consumed only by the check and published output.
shared #"EXTRACT Facility1 ResourceContractOverage" = let
    Source = #"IMPORT Facility1 Effort",
    Matches = Table.SelectRows(Source, each [Item] = "ResourceContractOverage" and [Kind] = "Table"),
    OverageTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Effort-All contract overage import", "Expected exactly one ResourceContractOverage table in the Facility1 Effort workbook.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Facility", "Facility-Abbrev", "Preferred Role", "Resource", "EmployeeID", "Name", "Employment Type",
        "Roster Start", "Roster End", "Roster Fortnights", "Contracted FN Hours", "Contracted Roster Hours", "Contracted Shift Equivalent",
        "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Allocated Roster Hours", "Allocated Shift Equivalent",
        "Over-Contract Hours", "Over-Contract Shift Equivalent", "Allocated Role List", "Limit Basis", "Is Over Contract", "Coverage Scope"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(OverageTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(OverageTable, RequiredColumns)
        else error Error.Record("Effort-All contract overage import", "ResourceContractOverage is missing required columns. Refresh Unit1 Effort first.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected, {{"Facility", type text}, {"Facility-Abbrev", type text}, {"Preferred Role", type text},
        {"Resource", Int64.Type}, {"EmployeeID", type text}, {"Name", type text}, {"Employment Type", type text},
        {"Roster Start", type date}, {"Roster End", type date}, {"Roster Fortnights", type number}, {"Contracted FN Hours", type number},
        {"Contracted Roster Hours", type number}, {"Contracted Shift Equivalent", type number}, {"Contracted Shifts", Int64.Type},
        {"Settings Max Availability", Int64.Type}, {"Effective Shift Cap", type number}, {"Allocated Roster Hours", type number},
        {"Allocated Shift Equivalent", type number}, {"Over-Contract Hours", type number}, {"Over-Contract Shift Equivalent", type number},
        {"Allocated Role List", type text}, {"Limit Basis", type text}, {"Is Over Contract", type logical}, {"Coverage Scope", type text}})
in
    Table.Buffer(Typed);

// Query: ResourceContractOverage_CHECK
// Purpose: Validate the Unit1/AIN scope, positive-overage rule and employee/resource/roster grain before publication.
// Notes: Keep this check connection-only; the published output blocks when any row fails.
shared ResourceContractOverage_CHECK = let
    Source = #"EXTRACT Facility1 ResourceContractOverage",
    Tolerance = 0.00000001,
    Facility1Unit = Text.AfterDelimiter(ResolveFacilityPath(Units{0}[Facility1]), "\", {0, RelativePosition.FromEnd}),
    DuplicateCount = Table.RowCount(Source) - Table.RowCount(Table.Distinct(Source, {"Facility", "EmployeeID", "Resource", "Roster Start", "Roster End"})),
    ScopeFailures = Table.RowCount(Table.SelectRows(Source, each [Facility] = null or [Facility] <> "Unit1"
        or [#"Preferred Role"] = null or [#"Preferred Role"] <> "AIN"
        or [#"Coverage Scope"] = null or [#"Coverage Scope"] <> "Unit1/AIN")),
    OverageFailures = Table.RowCount(Table.SelectRows(Source, each [#"Is Over Contract"] = null or [#"Is Over Contract"] <> true
        or [#"Over-Contract Hours"] = null or [#"Over-Contract Hours"] <= Tolerance)),
    Checks = {
        [Check = "Facility1 resolves to Unit1", Status = if Comparer.OrdinalIgnoreCase(Facility1Unit, "Unit1") = 0 then "Pass" else "Fail",
            Failures = if Comparer.OrdinalIgnoreCase(Facility1Unit, "Unit1") = 0 then 0 else 1,
            Details = if Comparer.OrdinalIgnoreCase(Facility1Unit, "Unit1") = 0 then null else "Facility1 resolves to " & Facility1Unit],
        [Check = "Unit1 AIN coverage only", Status = if ScopeFailures = 0 then "Pass" else "Fail", Failures = ScopeFailures, Details = null],
        [Check = "Positive exact-hour overages only", Status = if OverageFailures = 0 then "Pass" else "Fail", Failures = OverageFailures, Details = null],
        [Check = "Unique employee Resource roster keys", Status = if DuplicateCount = 0 then "Pass" else "Fail", Failures = DuplicateCount, Details = null]
    }
in
    Table.Buffer(Table.FromRecords(Checks, type table [Check = text, Status = text, Failures = number, Details = nullable text]));

// Query: ResourceContractOverage
// Purpose: Publish the validated Unit1/AIN contract overage table without appending it to date/shift facts.
// Output: Load this query to the worksheet table ResourceContractOverage.
shared ResourceContractOverage = let
    Failures = Table.SelectRows(ResourceContractOverage_CHECK, each true)
in
    Failures;

// Query: IMPORT Facility2 Effort
// Purpose: Share the second configured unit's Effort workbook across its table imports.
// Notes: Buffer the file binary once per evaluation; reuse the first import if both paths are equal.
shared #"IMPORT Facility2 Effort" = Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
    ResolveFacilityPath(Units{0}[Facility2]) & "\2. Calculations\Effort.xlsx")), null, true));

// Query: EXTRACT Facility2 EffortAllMatrixAG1_1D
// Purpose: Extract and date-align the second configured unit's effort matrix.
shared #"EXTRACT Facility2 EffortAllMatrixAG1_1D" = let
    Source = #"IMPORT Facility2 Effort",
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_Table,{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text},  {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"CapacityMaxHC", Int64.Type}, {"Allocation", type number}, {"DATESHIFT", type text}})
in
    #"Changed Type";

// Query: EXTRACT Facility2 RoleShiftAvailabilities
// Purpose: Extract the second configured unit's availability rows from the shared Effort import.
shared #"EXTRACT Facility2 RoleShiftAvailabilities" = let
    Source = #"IMPORT Facility2 Effort",
    RoleShiftAvailabilities_Table = Source{[Item="RoleShiftAvailabilities",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftAvailabilities_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"Capacity", type text},  {"Date", type date}, {"Shift", type text}})
in
    #"Changed Type1";

// Query: EXTRACT Facility2 ResShiftAllocation
// Purpose: Extract the second configured unit's resource allocation from the shared Effort import.
shared #"EXTRACT Facility2 ResShiftAllocation" = let
    Source = #"IMPORT Facility2 Effort",
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",null,43,Replacer.ReplaceValue,{"Period"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"Period", Int64.Type}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type1","1","2",Replacer.ReplaceText,{"Facility"})
in
    #"Replaced Value1";

// Query: Facility2 ResShiftAllocation1
// Purpose: Preserve the existing sorted Facility2 resource-allocation interface while reusing its single EXTRACT query.
shared #"Facility2 ResShiftAllocation1" = let
    Source = #"EXTRACT Facility2 ResShiftAllocation",
    #"Sorted Rows" = Table.Sort(Source,{{"Role", Order.Ascending}, {"Resource", Order.Ascending}, {"TimeDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    #"Sorted Rows";

shared ResShiftAllocation = let
    Source1 = #"EXTRACT Facility2 ResShiftAllocation",
    #"Appended Query" = Table.Combine({Source1, #"EXTRACT Facility1 ResShiftAllocation", #"EXTRACT Facility3 ResShiftAllocation"}),
    //CombinedTables = Table.Combine({Source1, #"EXTRACT Facility1 ResShiftAllocation"}),
    #"Renamed Columns" = Table.RenameColumns(#"Appended Query",{{"ShiftPeriod", "Shift"}, {"TimeDate", "Date"}, {"ResShiftFTE", "Effort"}, {"Type", "EffortType"}}),
    #"Filtered RES NULL" = Table.SelectRows(#"Renamed Columns", each ([Resource] <> null))
in
    #"Filtered RES NULL";

shared #"EffortAllMatrixAG1_1D-base !!" = let
   Source = #"EXTRACT Facility2 EffortAllMatrixAG1_1D",
    #"Appended Query" = Table.Combine({Source, #"EXTRACT Facility1 EffortAllMatrixAG1_1D", #"EXTRACT Facility3 EffortAllMatrixAG1_1D"}),
    #"Multiplied Column" = Table.TransformColumns(#"Appended Query", {{"Capacity", each _ * EffectiveAvailability, type number}}),
    #"Multiplied Column1" = Table.TransformColumns(#"Multiplied Column", {{"CapacityX", each _ * EffectiveAvailability, type number}}),
    //#"Appended Query" = Table.Combine({Source, #"EXTRACT Facility1 EffortAllMatrixAG1_1D"}),
    #"Renamed Columns" = Table.RenameColumns(#"Multiplied Column1",{{"Date", "DateX"}}),  // Add back Facility 2
    #"Added DATE ALIGNMENT" = Table.AddColumn(#"Renamed Columns", "Date", each if [Facility] = "ASHB" then
        Date.AddDays([DateX], -Units{0}[DateAlignment])
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
    Source = #"EXTRACT Facility2 RoleShiftAvailabilities",
    #"Appended FAC1" = Table.Combine({Source, #"EXTRACT Facility1 RoleShiftAvailabilities"}),
    //Table.Combine({#"EXTRACT Facility1 RoleShiftAvailabilities", #"EXTRACT Facility2 RoleShiftAvailabilities"}),    //return Facility 1
    #"Renamed Columns" = Table.RenameColumns(#"Appended FAC1",{{"Capacity", "EffortType"}, {"Availability", "Effort"}}),
    #"Multiplied Column" = Table.TransformColumns(#"Renamed Columns", {{"Effort", each _ * EffectiveAvailability, type number}}),
    #"Sorted Rows" = Table.Sort(#"Multiplied Column",{{"Facility", Order.Ascending}, {"Date", Order.Ascending}, {"Shift", Order.Ascending}, {"Resource", Order.Ascending}, {"EffortType", Order.Ascending}})
in
    #"Sorted Rows";

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
    #"Filtered NULL ROLE" = Table.SelectRows(Source, each ([Role] <> null)),
    #"Sorted Rows2" = Table.Sort(#"Filtered NULL ROLE",{{"Date", Order.Ascending}, {"EffortType", Order.Ascending}}),
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

shared #"LINK Unit Analysis" = let
    Source = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation="2.0", ViewMode="All"]),
    #"17ed8e30-5707-4af2-8e21-d68eb8184287" = Source{[Id="17ed8e30-5707-4af2-8e21-d68eb8184287"]}[Items],
    #"Removed Other Columns" = Table.SelectColumns(#"17ed8e30-5707-4af2-8e21-d68eb8184287",{"Title", "Effort Management Analysis"})
in
    #"Removed Other Columns";

// Query: INPUT EffortAllSettings
// Purpose: Read the two configured units and date alignment from one authoritative worksheet table.
// Inputs: Exactly one EffortAllSettings row with Facility1, Facility2 and DateAlignment columns.
shared Units = let
    Source = #"LINK Unit Analysis",
    #"Filtered UNITS" = Table.SelectRows(Source, each ([Effort Management Analysis] = true)),
    #"Added Index" = Table.AddIndexColumn(#"Filtered UNITS", "Index", 1, 1, Int64.Type),
    #"Added Prefix" = Table.TransformColumns(#"Added Index", {{"Index", each "Facility" & Text.From(_, "en-AU"), type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Added Prefix",{"Effort Management Analysis"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Columns", List.Distinct(#"Removed Columns"[Index]), "Index", "Title")
in
    #"Pivoted Column";

// Query: Folder
// Purpose: Preserve the scalar Folder interface as the resolved client root, without a trailing separator.
shared Folder = UnitL1PathTABLE{[#"Variable Name"="Client Path"]}[Value];

shared EffectiveAvailability = 7.6/8 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

shared Query1 = let
    Source = let
    Source = Excel.CurrentWorkbook(){[Name="EffortAllSettings"]}[Content],
    RequiredColumns = {"Facility1", "Facility2", "DateAlignment"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(Source)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(Source, RequiredColumns)
        else error Error.Record("Effort-All settings", "EffortAllSettings is missing required columns.", [MissingColumns = MissingColumns]),
    OneRow = if Table.RowCount(Selected) = 1 then Selected
        else error Error.Record("Effort-All settings", "EffortAllSettings must contain exactly one data row.", [Rows = Table.RowCount(Selected)]),
    Typed = Table.TransformColumnTypes(OneRow, {{"Facility1", type text}, {"Facility2", type text}, {"DateAlignment", Int64.Type}}),
    Trimmed = Table.TransformColumns(Typed, {{"Facility1", Text.Trim, type text}, {"Facility2", Text.Trim, type text}}),
    Settings = Trimmed{0},
    Validated = if Settings[Facility1] = null or Settings[Facility1] = ""
            or Settings[Facility2] = null or Settings[Facility2] = "" then
            error "EffortAllSettings Facility1 and Facility2 must be nonblank."
        else if Settings[DateAlignment] = null then error "EffortAllSettings DateAlignment must contain a whole number of days."
        else Trimmed
in
    Table.Buffer(Validated)
in
    Source;

shared #"IMPORT Facility3 Effort" = let
    Source = Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
    ResolveFacilityPath(Units{0}[Facility3]) & "\2. Calculations\Effort.xlsx")), null, true))
in
    Source;

shared #"EXTRACT Facility3 EffortAllMatrixAG1_1D" = let
    Source = #"IMPORT Facility3 Effort",
    EffortAllMatrixAG1_1D_Table = Source{[Item="EffortAllMatrixAG1_1D",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortAllMatrixAG1_1D_Table,{{"Facility", type text}, {"Role", type text}, {"Date", type date}, {"Shift", type text},  {"Demand", type number}, {"Capacity", type number}, {"CapacityX", type number}, {"CapacityMaxHC", Int64.Type}, {"Allocation", type number}, {"DATESHIFT", type text}})
in
    #"Changed Type";

shared #"EXTRACT Facility3 RoleShiftAvailabilities" = let
    Source = #"IMPORT Facility3 Effort",
    RoleShiftAvailabilities_Table = Source{[Item="RoleShiftAvailabilities",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(RoleShiftAvailabilities_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"Capacity", type text},  {"Date", type date}, {"Shift", type text}})
in
    #"Changed Type1";

shared #"EXTRACT Facility3 ResShiftAllocation" = let
    Source = #"IMPORT Facility3 Effort",
    ResShiftAllocation_Table = Source{[Item="ResShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Facility", type text}, {"Type", type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",null,43,Replacer.ReplaceValue,{"Period"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"Period", Int64.Type}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type1","1","2",Replacer.ReplaceText,{"Facility"})
in
    #"Replaced Value1";