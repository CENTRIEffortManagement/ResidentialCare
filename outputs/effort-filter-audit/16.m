section Section1;

// Query: EffortText
// Purpose: Normalize employee identifiers and descriptive text without converting identifiers to numbers.
shared EffortText = (value as any) as nullable text =>
    let Clean = try Text.Clean(Text.Trim(Text.From(value))) otherwise null
    in if Clean = "" then null else Clean;

// Query: EffortCheckResult
// Purpose: Convert a validation count or evaluation error into a consistent blocking check row.
shared EffortCheckResult = (name as text, evaluate as function) as record =>
    let Result = try evaluate()
    in [Check = name, Status = if Result[HasError] then "Fail" else if Result[Value] = 0 then "Pass" else "Fail",
        Failures = if Result[HasError] then null else Result[Value],
        Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null];

// Query: IMPORT CentriSyncPaths
// Purpose: Import the public-machine ResidentialCare path mappings once for the authoritative resolver.
shared #"IMPORT CentriSyncPaths" = let
    WorkbookNavigation = Excel.Workbook(Binary.Buffer(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx")), null, true),
    Matches = Table.SelectRows(WorkbookNavigation, each [Item] = "CentriSyncPaths" and [Kind] = "Table"),
    MappingTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Effort path resolver", "Expected exactly one CentriSyncPaths table in the public mapping workbook.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"SharepointRootUrl", "SyncedFolderRootPath"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(MappingTable)),
    Selected = if List.IsEmpty(MissingColumns) then MappingTable
        else error Error.Record("Effort path resolver", "CentriSyncPaths is missing required columns.", [MissingColumns = MissingColumns]),
    Typed = Table.TransformColumnTypes(Selected, {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}})
in
    Table.Buffer(Typed);

shared ResShiftAllocation = let
    Source = Table.NestedJoin(#"EXTRACT ResourceShiftAllocation", {"Name", "Role"}, #"IMPORT Table_Masterlist", {"Name", "Role"}, "IMPORT Table_Masterlist", JoinKind.LeftOuter),
    #"Expanded IMPORT Table_Masterlist" = Table.ExpandTableColumn(Source, "IMPORT Table_Masterlist", {"Name", "Resource"}, {"Name.1", "Resource"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Expanded IMPORT Table_Masterlist",{"ShiftPeriod", "Role", "TimeDate", "ResShiftFTE", "Resource"}),
    #"Merged Queries" = Table.NestedJoin(#"Removed Other Columns", {"TimeDate", "ShiftPeriod"}, DateShiftPeriod, {"Date", "Shift"}, "DateShiftPeriod", JoinKind.LeftOuter),
    #"Expanded DateShiftPeriod" = Table.ExpandTableColumn(#"Merged Queries", "DateShiftPeriod", {"Period"}, {"Period"}),
    #"Added Custom" = Table.AddColumn(#"Expanded DateShiftPeriod", "Facility", each Unit),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "Type", each "Allocation")
in
    #"Added Custom1";

shared RoleList = let
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WCvJTitWJVnIOdlWKjQUA", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"Column1", "RoleList"}}),
    RoleList1 = #"Renamed Columns"[RoleList]
in
    RoleList1;

// Query: IMPORT Demand Workbook
// Purpose: Import Demand.xlsx once for the existing demand extraction.
shared #"IMPORT Demand Workbook" = let
    WorkbookBinary = Binary.Buffer(File.Contents(Unit1Path & "\2. Calculations\Demand.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in
    Table.Buffer(WorkbookNavigation);

// Query: IMPORT DemandCorrected
// Purpose: Extract the existing demand table from the shared Demand workbook import.
shared #"IMPORT DemandCorrected" = let
    Source = #"IMPORT Demand Workbook",
    ShiftDemandHCAverageANACC_Table = Source{[Item="ShiftDemandHCAverageANACC",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ShiftDemandHCAverageANACC_Table,{{"Facility", type any}})
in
    Table.Buffer(#"Changed Type");

// Query: RoleDATEShiftDemand
// Purpose: Preserve the established workbook interface used by existing Excel connections.
// Notes: Compatibility reference only; the Demand workbook is opened by IMPORT Demand Workbook.
shared RoleDATEShiftDemand = #"IMPORT DemandCorrected";

// Query: IMPORT Capacity Workbook
// Purpose: Import Capacity.xlsx once for all availability consumers.
shared #"IMPORT Capacity Workbook" = let
    WorkbookBinary = Binary.Buffer(File.Contents(Unit1Path & "\2. Calculations\Capacity.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in
    Table.Buffer(WorkbookNavigation);

// Query: IMPORT AvailabilityDeveloped
// Purpose: Extract the AvailabilityDeveloped table from the shared Capacity workbook import.
shared #"IMPORT AvailabilityDeveloped" = let
    Source = #"IMPORT Capacity Workbook",
    AvailabilityDeveloped_Table = Source{[Item="AvailabilityDeveloped",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(AvailabilityDeveloped_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"AvailabilityType", type text}, {"ResAvailability", Int64.Type}, {"ResMaxAvail", Int64.Type}, {"Facility", type text}})
in
    Table.Buffer(#"Changed Type");

// Query: AvailabilityDeveloped2
// Purpose: Preserve the established workbook interface used by existing Excel connections.
// Notes: Compatibility reference only; Capacity.xlsx is opened once by IMPORT Capacity Workbook.
shared AvailabilityDeveloped2 = #"IMPORT AvailabilityDeveloped";

// Query: IMPORT Allocation
// Purpose: Import and buffer Allocation.xlsx once for its Allocation and ResourceShiftAllocation extractions.
// Output: Workbook navigation consumed only by the two EXTRACT queries below.
shared #"IMPORT Allocation" = let
    WorkbookBinary = Binary.Buffer(File.Contents(Unit1Path & "\2. Calculations\Allocation.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in
    Table.Buffer(WorkbookNavigation);

// Query: IMPORT Allocation Workbook
// Purpose: Preserve the already-synced interim workbook interface so its Excel connection is not orphaned.
// Notes: Compatibility reference only; IMPORT Allocation is the single buffered external read.
shared #"IMPORT Allocation Workbook" = #"IMPORT Allocation";

// Query: EXTRACT Allocation
// Purpose: Extract the first of two tables from the consolidated Allocation workbook import.
shared #"EXTRACT Allocation" = let
    Source = #"IMPORT Allocation",
    RoleShiftAllocation_Table = Source{[Item="Allocation",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RoleShiftAllocation_Table,{{"ShiftDate", type date}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{{"ShiftDate", "Date"}})
in
    Table.Buffer(#"Renamed Columns");

// Query: IMPORT Settings Data
// Purpose: Import Settings Data.xlsx once for all existing and contract-reporting settings.
shared #"IMPORT Settings Data" = let
    WorkbookBinary = Binary.Buffer(File.Contents(Unit1Path & "\2. Calculations\Settings Data.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in
    Table.Buffer(WorkbookNavigation);

// Query: EXTRACT PermutationDimensions
// Purpose: Extract the roster permutation table from the consolidated Settings Data import.
shared #"EXTRACT PermutationDimensions" = let
    Source = #"IMPORT Settings Data",
    PermutationDimensions_Table = Source{[Item="PermutationDimensions",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(PermutationDimensions_Table,{{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    Table.Buffer(#"Changed Type");

// Query: IMPORT PermutationDimensions
// Purpose: Preserve the established workbook interface used by existing Excel connections.
// Notes: Compatibility reference only; Settings Data.xlsx is opened once by IMPORT Settings Data.
shared #"IMPORT PermutationDimensions" = #"EXTRACT PermutationDimensions";

// Query: EXTRACT MaxCapacity
// Purpose: Extract the MaxCapacityFactor scalar from the consolidated Settings Data import.
shared #"EXTRACT MaxCapacity" = let
    Source = #"IMPORT Settings Data",
    MaxCapacity_Table = Source{[Item="MaxCapacity",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(MaxCapacity_Table,{{"MaxCapacityFactor", Int64.Type}}),
    MaxCapacityFactor = #"Changed Type"{0}[MaxCapacityFactor]
in
    MaxCapacityFactor;

// Query: IMPORTMaxCapacity
// Purpose: Preserve the established scalar interface used by existing Excel connections and calculations.
// Notes: Compatibility reference only; the value is extracted once by EXTRACT MaxCapacity.
shared IMPORTMaxCapacity = #"EXTRACT MaxCapacity";

shared RoleShiftAllocationDay = let
    Source = #"EXTRACT Allocation",
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"ShiftPeriod", type text}, {"Role", type text}})
in
    #"Changed Type";

shared RoleShiftCapacity = let
    Source = #"IMPORT AvailabilityDeveloped",
    #"Filtered Rows" = Table.SelectRows(Source, each ([AvailabilityType] = "C###")),
    #"Renamed Columns1" = Table.RenameColumns(#"Filtered Rows",{{"AvailabilityType", "Capacity"}}),
    #"Merged Queries" = Table.NestedJoin(#"Renamed Columns1", {"Period"}, DateShiftPeriod, {"Period"}, "IMPORT PermutationDimensions (2)", JoinKind.LeftOuter),
    #"Expanded PERMUTATIONS DATE SHIFT" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT PermutationDimensions (2)", {"Date", "Shift"}, {"Date", "Shift"}),
    #"Grouped CAPACITY CAPCAITY MAX" = Table.Group(#"Expanded PERMUTATIONS DATE SHIFT", {"Role", "Date", "Shift"}, {{"RoleShiftCapacity", each List.Sum([Availability]), type nullable number}, {"RoleMaxCapacity", each List.Max([ResMaxAvail]), type nullable number}}),
    #"Inserted Day Name" = Table.AddColumn(#"Grouped CAPACITY CAPCAITY MAX", "Day Name", each Date.DayOfWeekName([Date]), type text),
    #"Inserted First Characters" = Table.AddColumn(#"Inserted Day Name", "First Characters", each Text.Start([Day Name], 3), type text),
    #"Renamed Columns" = Table.RenameColumns(#"Inserted First Characters",{{"First Characters", "Day"}}),
    #"Removed Columns" = Table.RemoveColumns(#"Renamed Columns",{"Day Name"}),
    #"Added CAPACITYX" = Table.AddColumn(#"Removed Columns", "CapacityX", each [RoleShiftCapacity] * #"EXTRACT MaxCapacity"),
    #"Sorted Rows" = Table.Sort(#"Added CAPACITYX",{{"Date", Order.Ascending}, {"Role", Order.Ascending}})
in
    #"Sorted Rows";

shared #"EffortAllMatrixAG1-1D" = let
    Source = Table.NestedJoin(RoleShiftCapacity, {"Date", "Role", "Shift"}, #"IMPORT DemandCorrected", {"Date", "Role", "ShiftPeriod"}, "Demand", JoinKind.FullOuter),
    #"Expanded DEMAND" = Table.ExpandTableColumn(Source, "Demand", {"Date", "Role", "ShiftDemandHCAverage", "ShiftPeriod"}, {"Date.1", "Role.1", "ShiftDemandHCAverage", "ShiftPeriod"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded DEMAND",{{"Date.1", type date}}),
    #"Renamed Columns2" = Table.RenameColumns(#"Changed Type",{{"Date", "Date-Capacity"}, {"Role", "Role-Capacity"}, {"Role.1", "Role"}, {"Shift", "Shift-Capacity"}, {"Date.1", "Date"}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Renamed Columns2",{"Role-Capacity", "Date-Capacity", "Shift-Capacity"}),
    #"Merged ALLOCATION" = Table.NestedJoin(#"Removed Columns1", {"Role", "Date", "ShiftPeriod"}, #"EXTRACT Allocation", {"Role", "Date", "ShiftPeriod"}, "RoleShiftAllocation", JoinKind.LeftOuter),
    #"Expanded RoleShiftAllocation" = Table.ExpandTableColumn(#"Merged ALLOCATION", "RoleShiftAllocation", {"RoleShiftFTE"}, {"RoleShiftFTE"}),
    #"Sorted Rows" = Table.Sort(#"Expanded RoleShiftAllocation",{{"Date", Order.Ascending}, {"Role", Order.Ascending}}),
    #"Renamed Columns" = Table.RenameColumns(#"Sorted Rows",{{"RoleShiftCapacity", "Capacity"}, {"ShiftDemandHCAverage", "Demand"}, {"RoleMaxCapacity", "CapacityMaxHC"}, {"RoleShiftFTE", "Allocation"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Role", "Date", "ShiftPeriod", "Day", "Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation"}),
    #"Replaced Value" = Table.ReplaceValue(#"Reordered Columns",null,0,Replacer.ReplaceValue,{"Allocation"}),
    #"Removed Columns" = Table.RemoveColumns(#"Replaced Value",{"Day"}),
    #"Inserted Merged Column" = Table.AddColumn(#"Removed Columns", "Merged", each Text.Combine({Text.From([Date], "en-AU"), [ShiftPeriod]}, ""), type text),
    #"Filtered Rows" = Table.SelectRows(#"Inserted Merged Column", each ([Role] <> "Endorsed Enrolled Nurse")),
    #"Renamed Columns1" = Table.RenameColumns(#"Filtered Rows",{{"Merged", "DATESHIFT"}, {"ShiftPeriod", "Shift"}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Renamed Columns1",null,0,Replacer.ReplaceValue,{"Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation"}),
    #"Added Facility" = Table.AddColumn(#"Replaced Value1", "Facility", each Unit),
    #"Reordered Columns1" = Table.ReorderColumns(#"Added Facility",{"Facility", "Role", "Date", "Shift", "Demand", "Capacity", "CapacityMaxHC", "CapacityX", "Allocation", "DATESHIFT"}),
    #"Sorted Rows1" = Table.Sort(#"Reordered Columns1",{{"Role", Order.Ascending}, {"DATESHIFT", Order.Ascending}})
in
    #"Sorted Rows1";

shared #"EffortDAMatrixAG1-1D" = let
    Source = Table.NestedJoin(RoleShiftAllocationDay, {"Date", "ShiftPeriod", "Role"}, #"IMPORT DemandCorrected", {"Date", "ShiftPeriod", "Role"}, "RoleDATEShiftDemand", JoinKind.LeftOuter),
    #"Expanded RoleDATEShiftDemand" = Table.ExpandTableColumn(Source, "RoleDATEShiftDemand", {"ShiftDemandHCAverage"}, {"ShiftDemandHCAverage"})
in
    #"Expanded RoleDATEShiftDemand";

shared RoleShiftAvailabilities = let
    Source = #"IMPORT AvailabilityDeveloped",
    #"Renamed Columns1" = Table.RenameColumns(Source,{{"AvailabilityType", "Capacity"}}),
    #"Merged Queries" = Table.NestedJoin(#"Renamed Columns1", {"Period"}, DateShiftPeriod, {"Period"}, "IMPORT PermutationDimensions (2)", JoinKind.LeftOuter),
    #"Expanded PERMUTATIONS DATE SHIFT" = Table.ExpandTableColumn(#"Merged Queries", "IMPORT PermutationDimensions (2)", {"Date", "Shift"}, {"Date", "Shift"})
in
    #"Expanded PERMUTATIONS DATE SHIFT";

shared DateShiftPeriod = let
    Source = #"EXTRACT PermutationDimensions",
    #"Removed Columns" = Table.RemoveColumns(Source,{"RolesList", "Day"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Columns"),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Duplicates",{{"Shifts", "Shift"}})
in
    #"Renamed Columns";

// Query: EXTRACT ResourceShiftAllocation
// Purpose: Extract the second table from the consolidated Allocation workbook import.
shared #"EXTRACT ResourceShiftAllocation" = let
    Source = #"IMPORT Allocation",
    ResourceShiftAllocation_Table = Source{[Item="ResourceShiftAllocation",Kind="Table"]}[Data],
    #"Changed Type1" = Table.TransformColumnTypes(ResourceShiftAllocation_Table,{{"ShiftPeriod", type text}, {"Name", type text}, {"Role", type text}, {"TimeDate", type date}, {"ResShiftFTE", type number}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"Role", Order.Ascending}, {"Name", Order.Ascending}, {"TimeDate", Order.Ascending}, {"ShiftPeriod", Order.Ascending}})
in
    #"Sorted Rows";

// Query: IMPORT ResourceShiftAllocation
// Purpose: Preserve the established workbook interface used by existing Excel connections.
// Notes: Compatibility reference only; Allocation.xlsx is opened once by IMPORT Allocation.
shared #"IMPORT ResourceShiftAllocation" = #"EXTRACT ResourceShiftAllocation";

// Query: IMPORT StaffListMaster
// Purpose: Import StaffListMaster.xlsx once for its extended Table_Masterlist interface.
shared #"IMPORT StaffListMaster" = let
    WorkbookBinary = Binary.Buffer(File.Contents(Unit1Path & "\2. Calculations\StaffListMaster.xlsx")),
    WorkbookNavigation = Excel.Workbook(WorkbookBinary, null, true)
in
    Table.Buffer(WorkbookNavigation);

// Query: IMPORT Table_Masterlist
// Purpose: Extract the resource master, employee contracts and effective capacity caps.
shared #"IMPORT Table_Masterlist" = let
    Source = #"IMPORT StaffListMaster",
    Table_Masterlist_Table = Source{[Item="Table_Masterlist",Kind="Table"]}[Data],
    RequiredColumns = {"Name", "Role", "Resource", "Source", "EmployeeID", "Facility-Abbrev", "EmploymentType", "PreferredRole",
        "Worker Record Status", "Worker Contract Issue", "Contracted FN Hours", "Roster Start", "Roster End", "Roster Fortnights",
        "Contracted Roster Hours", "Contracted Shift Equivalent", "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(Table_Masterlist_Table)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(Table_Masterlist_Table, RequiredColumns)
        else error Error.Record("Effort staff master import", "Table_Masterlist is missing required contract columns. Refresh StaffListMaster first.", [MissingColumns = MissingColumns]),
    NormalizedText = Table.TransformColumns(Selected, {
        {"Name", EffortText, type nullable text}, {"Role", EffortText, type nullable text}, {"Source", EffortText, type nullable text},
        {"EmployeeID", EffortText, type nullable text}, {"Facility-Abbrev", EffortText, type nullable text},
        {"EmploymentType", EffortText, type nullable text}, {"PreferredRole", EffortText, type nullable text},
        {"Worker Record Status", EffortText, type nullable text}, {"Worker Contract Issue", EffortText, type nullable text}, {"Limit Basis", EffortText, type nullable text}}),
    #"Changed Type" = Table.TransformColumnTypes(NormalizedText, {{"Resource", Int64.Type}, {"Contracted FN Hours", type nullable number},
        {"Roster Start", type date}, {"Roster End", type date}, {"Roster Fortnights", type number},
        {"Contracted Roster Hours", type nullable number}, {"Contracted Shift Equivalent", type nullable number},
        {"Contracted Shifts", Int64.Type}, {"Settings Max Availability", Int64.Type}, {"Effective Shift Cap", type nullable number}})
in
    Table.Buffer(#"Changed Type");

// Query: UnitL1PathTABLE
// Purpose: Resolve this exact Effort workbook through a boundary-safe longest-prefix CentriSyncPaths match.
// Output: Existing Variable Name / Value interface; Root Path remains the 2. Calculations folder.
shared UnitL1PathTABLE = // Version 25.02 flexible ResidentialCare
let
    NormalizePath = (value as nullable text) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(value), "/", "\"), "\"),
    PathInputs = Table.SelectRows(Excel.CurrentWorkbook(), each Comparer.OrdinalIgnoreCase([Name], "FilePathUrl") = 0),
    PathSource = if Table.RowCount(PathInputs) = 1 then PathInputs{0}[Content]
        else error "Expected exactly one FilePathUrl table or named cell in Effort.xlsx.",
    PathColumn = if Table.HasColumns(PathSource, {"FilePath"}) then Table.SelectColumns(PathSource, {"FilePath"})
        else if Table.ColumnNames(PathSource) = {"Column1"} then Table.RenameColumns(PathSource, {{"Column1", "FilePath"}})
        else error "FilePathUrl must contain a FilePath column or be a single named cell.",
    SinglePath = if Table.RowCount(PathColumn) = 1 then PathColumn{0}[FilePath]
        else error "FilePathUrl must contain exactly one data row.",
    RawFilePath = NormalizePath(if SinglePath = null then null else Text.From(SinglePath)),
    NonBlankFilePath = if RawFilePath = null or RawFilePath = "" then
        error "FilePathUrl is blank. Save Effort.xlsx and recalculate its CELL filename formula."
        else RawFilePath,
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet; retain only the workbook path.
    WorkbookPath = if Text.Contains(NonBlankFilePath, "[") and Text.Contains(NonBlankFilePath, "]") then
        Text.BeforeDelimiter(NonBlankFilePath, "[") & Text.BetweenDelimiters(NonBlankFilePath, "[", "]")
        else NonBlankFilePath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    FilePath = if Comparer.OrdinalIgnoreCase(InputFileName, "Effort.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl identifies another workbook. Save Effort.xlsx and recalculate its CELL filename formula.",
    NormalizedMappings = Table.TransformColumns(#"IMPORT CentriSyncPaths",
        {{"SharepointRootUrl", NormalizePath, type nullable text}, {"SyncedFolderRootPath", NormalizePath, type nullable text}}),
    SharePointCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    SharePointDocumentsCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot",
        each if [SharepointRootUrl] = null or [SharepointRootUrl] = "" then null else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalCandidates = Table.AddColumn(NormalizedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    MatchCandidates = Table.AddColumn(Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
        "MatchRootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(MatchCandidates, each [MatchRootLength] > 0
        and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
        and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    LongestMatches = if Table.IsEmpty(SortedMatches) then error "FilePathUrl has no usable CentriSyncPaths mapping: " & FilePath
        else Table.SelectRows(SortedMatches, each [MatchRootLength] = SortedMatches{0}[MatchRootLength]),
    Destinations = List.Distinct(LongestMatches[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    BestMatch = if List.Count(Destinations) = 1 then LongestMatches{0}
        else error "CentriSyncPaths contains conflicting local folders for Effort.xlsx.",
    LocalRoot = BestMatch[SyncedFolderRootPath],
    ValidLocalRoot = if (Text.Length(LocalRoot) >= 3 and Text.Range(LocalRoot, 1, 2) = ":\") or Text.StartsWith(LocalRoot, "\\") then LocalRoot
        else error "CentriSyncPaths SyncedFolderRootPath must be an absolute local or UNC path.",
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    LocalFullPath = ValidLocalRoot & RelativePath,
    RootPath = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    CalculationSuffix = "\2. Calculations",
    UnitRoot = if Text.EndsWith(RootPath, CalculationSuffix, Comparer.OrdinalIgnoreCase) then
        Text.Start(RootPath, Text.Length(RootPath) - Text.Length(CalculationSuffix))
        else error "Effort.xlsx must be saved in a Unit 2. Calculations folder.",
    Segments = List.Select(Text.Split(UnitRoot, "\"), each _ <> ""),
    ResidentialCareIndex = List.PositionOf(Segments, "ResidentialCare", Occurrence.First, Comparer.OrdinalIgnoreCase),
    UnitsIndex = List.PositionOf(Segments, "UNITS", Occurrence.Last, Comparer.OrdinalIgnoreCase),
    ValidUnit = if UnitsIndex >= 0 and List.Count(Segments) = UnitsIndex + 2 then Segments{UnitsIndex + 1}
        else error "Resolved Effort path must be under ResidentialCare\<client>\UNITS\<unit>\2. Calculations.",
    UserName = try Text.BeforeDelimiter(Text.AfterDelimiter(RootPath, "C:\Users\"), "\") otherwise null,
    Client = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 1 then Segments{ResidentialCareIndex + 1} else null,
    DateValue = if ResidentialCareIndex >= 0 and List.Count(Segments) > ResidentialCareIndex + 2 then Segments{ResidentialCareIndex + 2} else null,
    TABLE = #table({"Variable Name", "Value"}, {
        {"UserName", UserName}, {"Root Path", RootPath}, {"Client", Client}, {"Date", DateValue},
        {"Unit", ValidUnit}, {"FileName", InputFileName}})
in
    Table.Buffer(TABLE);

// Query: Unit1Path
// Purpose: Preserve the existing scalar Unit root derived from the validated Effort workbook path.
shared Unit1Path = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Root Path")),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Variable Name"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Value", "Folder"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns","\2. Calculations","",Replacer.ReplaceText,{"Folder"}),
    Folder = #"Replaced Value"{0}[Folder]
in
    Folder;

// Query: Unit
// Purpose: Preserve the existing dynamic unit label derived from the validated workbook path.
shared Unit = let
    Source = UnitL1PathTABLE,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Variable Name] = "Unit")),
    Value = #"Filtered Rows"{0}[Value]
in
    Value;

// Query: FilePath-1Input
// Purpose: Derive the Unit1 input folder from the resolved 2. Calculations folder.
shared #"FilePath-1Input" = let
    CalculationsFolder = UnitL1PathTABLE{[Variable Name = "Root Path"]}[Value],
    Suffix = "\2. Calculations",
    UnitFolder = if Text.EndsWith(CalculationsFolder, Suffix, Comparer.OrdinalIgnoreCase) then
        Text.Start(CalculationsFolder, Text.Length(CalculationsFolder) - Text.Length(Suffix))
        else error "Resolved Effort folder is not the expected Unit 2. Calculations folder."
in
    UnitFolder & "\1. Input";

// Query: EXTRACT Effort RosterDates
// Purpose: Validate the distinct contiguous Settings roster dates used to restrict allocation hours.
shared #"EXTRACT Effort RosterDates" = let
    Selected = Table.SelectColumns(#"EXTRACT PermutationDimensions", {"Date"}),
    MissingDates = Table.RowCount(Table.SelectRows(Selected, each [Date] = null)),
    DistinctDates = Table.Sort(Table.Distinct(Table.SelectRows(Selected, each [Date] <> null)), {{"Date", Order.Ascending}}),
    DateCount = Table.RowCount(DistinctDates),
    SpanDays = if DateCount = 0 then 0 else Duration.Days(List.Max(DistinctDates[Date]) - List.Min(DistinctDates[Date])) + 1,
    Validated = if MissingDates > 0 then error "PermutationDimensions contains blank roster dates."
        else if DateCount = 0 then error "PermutationDimensions contains no roster dates."
        else if DateCount <> SpanDays then error "Settings roster dates must be contiguous."
        else if Number.Mod(DateCount, 14) <> 0 then error "Settings roster date count must be a whole number of fortnights."
        else DistinctDates
in
    Table.Buffer(Validated);

// Query: EXTRACT Effort ShiftDuration
// Purpose: Validate the standard shift duration in hours for exact shift-equivalent reporting.
shared #"EXTRACT Effort ShiftDuration" = let
    Matches = Table.SelectRows(#"IMPORT Settings Data", each [Item] = "ShiftDuration" and List.Contains({"Table", "DefinedName"}, [Kind])),
    SettingsTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Effort shift duration", "Expected exactly one ShiftDuration table or defined name.", [Matches = Table.RowCount(Matches)]),
    WithHeaders = if Table.HasColumns(SettingsTable, "ShiftDuration") then SettingsTable else Table.PromoteHeaders(SettingsTable, [PromoteAllScalars = true]),
    Values = Table.Column(WithHeaders, "ShiftDuration"),
    ShiftDuration = if List.Count(Values) = 1 then Number.From(Values{0}) else error "ShiftDuration must contain exactly one value.",
    Validated = if ShiftDuration = null or Number.IsNaN(ShiftDuration) or Number.Abs(ShiftDuration) = #infinity or ShiftDuration <= 0 then
        error "ShiftDuration must be one positive finite number in hours." else ShiftDuration
in
    Validated;

// Query: IMPORT AllocationExtracted
// Purpose: Import the Unit1 AllocationExtracted table used for exact employee roster-hour reporting.
// Notes: This is distinct from Allocation.xlsx because actual Hours and employee Code are required.
shared #"IMPORT AllocationExtracted" = let
    WorkbookNavigation = Excel.Workbook(Binary.Buffer(File.Contents(#"FilePath-1Input" & "\1-AllocationExtracted.xlsx")), null, true),
    Matches = Table.SelectRows(WorkbookNavigation, each [Item] = "AllocationExtracted" and [Kind] = "Table"),
    AllocationTable = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error Error.Record("Effort actual-hours import", "Expected exactly one AllocationExtracted table in 1-AllocationExtracted.xlsx.", [Matches = Table.RowCount(Matches)]),
    RequiredColumns = {"Code", "Name", "Date", "Hours", "Role"},
    MissingColumns = List.Difference(RequiredColumns, Table.ColumnNames(AllocationTable)),
    Selected = if List.IsEmpty(MissingColumns) then Table.SelectColumns(AllocationTable, RequiredColumns)
        else error Error.Record("Effort actual-hours import", "AllocationExtracted is missing required columns.", [MissingColumns = MissingColumns])
in
    Table.Buffer(Selected);

// Query: AllocationContractHours_Prepare
// Purpose: Parse actual employee allocation hours, restrict them to Settings roster dates and aggregate all roles once by employee.
// Output: A record containing source diagnostics, roster rows and one total row per employee.
shared AllocationContractHours_Prepare = let
    RosterDates = #"EXTRACT Effort RosterDates"[Date],
    WithSourceRow = Table.AddIndexColumn(#"IMPORT AllocationExtracted", "Source Row", 1, 1, Int64.Type),
    WithPrepared = Table.AddColumn(WithSourceRow, "Prepared", each
        let
            EmployeeID = EffortText([Code]),
            EmployeeName = EffortText([Name]),
            AllocationRole = EffortText([Role]),
            DateAttempt = try Date.From([Date]),
            HoursAttempt = try Number.From([Hours]),
            AllocationDate = if DateAttempt[HasError] then null else DateAttempt[Value],
            AllocatedHours = if HoursAttempt[HasError] then null else HoursAttempt[Value],
            Issue = if EmployeeID = null then "Missing employee Code"
                else if EmployeeName = null then "Missing employee Name"
                else if AllocationRole = null then "Missing allocated Role"
                else if DateAttempt[HasError] or AllocationDate = null then "Invalid allocation Date"
                else if HoursAttempt[HasError] or AllocatedHours = null then "Invalid allocation Hours"
                else if Number.IsNaN(AllocatedHours) or Number.Abs(AllocatedHours) = #infinity then "Non-finite allocation Hours"
                else if AllocatedHours < 0 then "Negative allocation Hours"
                else null
        in [EmployeeID = EmployeeID, EmployeeName = EmployeeName, AllocationDate = AllocationDate,
            AllocatedHours = AllocatedHours, AllocationRole = AllocationRole, Issue = Issue],
        type [EmployeeID = nullable text, EmployeeName = nullable text, AllocationDate = nullable date,
            AllocatedHours = nullable number, AllocationRole = nullable text, Issue = nullable text]),
    Expanded = Table.ExpandRecordColumn(WithPrepared, "Prepared",
        {"EmployeeID", "EmployeeName", "AllocationDate", "AllocatedHours", "AllocationRole", "Issue"}),
    PreparedRows = Table.SelectColumns(Expanded, {"Source Row", "EmployeeID", "EmployeeName", "AllocationDate", "AllocatedHours", "AllocationRole", "Issue"}),
    RosterRows = Table.SelectRows(PreparedRows, each [AllocationDate] <> null and List.Contains(RosterDates, [AllocationDate])),
    ValidRosterRows = Table.SelectRows(RosterRows, each [Issue] = null),
    Grouped = Table.Group(ValidRosterRows, {"EmployeeID"}, {
        {"Allocated Roster Hours", each List.Sum([AllocatedHours]), type number},
        {"Allocated Role List", each Text.Combine(List.Sort(List.Distinct(List.RemoveNulls([AllocationRole]))), ", "), type text},
        {"Allocation Row Count", each Table.RowCount(_), Int64.Type}})
in
    [AllPreparedRows = Table.Buffer(PreparedRows), RosterRows = Table.Buffer(RosterRows), EmployeeTotals = Table.Buffer(Grouped)];

// Query: AINContractResources_Prepare
// Purpose: Select the preferred AIN employee/resource contract at Resource grain from Table_Masterlist.
// Notes: Allocation-only Resources do not enter this population because they are not preferred AIN availability Resources.
shared AINContractResources_Prepare = let
    Filtered = Table.SelectRows(#"IMPORT Table_Masterlist", each [Source] = "Availability" and [Role] = "AIN" and [PreferredRole] = "AIN"),
    Selected = Table.SelectColumns(Filtered, {"EmployeeID", "Name", "Resource", "Facility-Abbrev", "EmploymentType", "PreferredRole",
        "Worker Record Status", "Worker Contract Issue", "Roster Start", "Roster End", "Roster Fortnights", "Contracted FN Hours",
        "Contracted Roster Hours", "Contracted Shift Equivalent", "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Limit Basis"})
in
    Table.Buffer(Selected);

// Query: ResourceContractOverage_Prepare
// Purpose: Compare exact roster allocation hours with exact contracted roster hours for preferred AIN employees.
// Notes: Floored Contracted Shifts and Effective Shift Cap are context only; neither determines an hours breach.
shared ResourceContractOverage_Prepare = let
    ShiftDuration = #"EXTRACT Effort ShiftDuration",
    Tolerance = 0.00000001,
    AllocationTotals = AllocationContractHours_Prepare[EmployeeTotals],
    JoinedAllocation = Table.NestedJoin(AINContractResources_Prepare, {"EmployeeID"}, AllocationTotals, {"EmployeeID"}, "Allocation", JoinKind.LeftOuter),
    ExpandedAllocation = Table.ExpandTableColumn(JoinedAllocation, "Allocation",
        {"Allocated Roster Hours", "Allocated Role List", "Allocation Row Count"},
        {"Allocated Roster Hours", "Allocated Role List", "Allocation Row Count"}),
    FilledHours = Table.ReplaceValue(ExpandedAllocation, null, 0, Replacer.ReplaceValue, {"Allocated Roster Hours", "Allocation Row Count"}),
    AddedAllocatedEquivalent = Table.AddColumn(FilledHours, "Allocated Shift Equivalent", each [#"Allocated Roster Hours"] / ShiftDuration, type number),
    AddedOverHours = Table.AddColumn(AddedAllocatedEquivalent, "Over-Contract Hours", each
        if [#"Contracted Roster Hours"] = null then null else [#"Allocated Roster Hours"] - [#"Contracted Roster Hours"], type nullable number),
    AddedOverEquivalent = Table.AddColumn(AddedOverHours, "Over-Contract Shift Equivalent", each
        if [#"Over-Contract Hours"] = null then null else [#"Over-Contract Hours"] / ShiftDuration, type nullable number),
    AddedIsOver = Table.AddColumn(AddedOverEquivalent, "Is Over Contract", each
        let IsCasualFallback = [#"Limit Basis"] <> null and Text.StartsWith([#"Limit Basis"], "Casual fallback", Comparer.OrdinalIgnoreCase)
        in not IsCasualFallback and [#"Over-Contract Hours"] <> null and [#"Over-Contract Hours"] > Tolerance, type logical),
    AddedFacility = Table.AddColumn(AddedIsOver, "Facility", each Unit, type text),
    AddedScope = Table.AddColumn(AddedFacility, "Coverage Scope", each "Unit1/AIN", type text),
    Renamed = Table.RenameColumns(AddedScope, {{"PreferredRole", "Preferred Role"}, {"EmploymentType", "Employment Type"}}),
    Ordered = Table.ReorderColumns(Renamed, {"Facility", "Facility-Abbrev", "Preferred Role", "Resource", "EmployeeID", "Name", "Employment Type",
        "Roster Start", "Roster End", "Roster Fortnights", "Contracted FN Hours", "Contracted Roster Hours", "Contracted Shift Equivalent",
        "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Allocated Roster Hours", "Allocated Shift Equivalent",
        "Over-Contract Hours", "Over-Contract Shift Equivalent", "Allocated Role List", "Limit Basis", "Is Over Contract", "Coverage Scope",
        "Worker Record Status", "Worker Contract Issue", "Allocation Row Count"})
in
    Table.Buffer(Ordered);

// Query: ResourceContractOverage_CHECK
// Purpose: Validate exact-hours aggregation, AIN contract coverage, roster alignment and the one-row output grain.
// Notes: Casual fallback employees are diagnostic and are deliberately excluded from over-contract publication.
shared ResourceContractOverage_CHECK = let
    Tolerance = 0.00000001,
    RosterDates = #"EXTRACT Effort RosterDates",
    RosterStart = List.Min(RosterDates[Date]),
    RosterEnd = List.Max(RosterDates[Date]),
    RosterFortnights = Number.From(Table.RowCount(RosterDates)) / 14,
    ShiftDuration = #"EXTRACT Effort ShiftDuration",
    Contracts = AINContractResources_Prepare,
    AllocationPrepared = AllocationContractHours_Prepare[AllPreparedRows],
    AllocationRosterRows = AllocationContractHours_Prepare[RosterRows],
    AllocationTotals = AllocationContractHours_Prepare[EmployeeTotals],
    Prepared = ResourceContractOverage_Prepare,
    CasualFallbackRows = Table.SelectRows(Contracts, each [#"Limit Basis"] <> null and Text.StartsWith([#"Limit Basis"], "Casual fallback", Comparer.OrdinalIgnoreCase)),
    PublishedRows = Table.SelectRows(Prepared, each [#"Is Over Contract"]),
    BlockingChecks = {
        EffortCheckResult("Effort workbook resolves to Unit1", () => if Comparer.OrdinalIgnoreCase(Unit, "Unit1") = 0 then 0 else 1),
        EffortCheckResult("Valid roster dates and shift duration", () => if Table.RowCount(RosterDates) > 0 and RosterFortnights > 0 and ShiftDuration > 0 then 0 else 1),
        EffortCheckResult("Valid AllocationExtracted dates", () => Table.RowCount(Table.SelectRows(AllocationPrepared, each [AllocationDate] = null))),
        EffortCheckResult("Valid roster allocation rows", () => Table.RowCount(Table.SelectRows(AllocationRosterRows, each [Issue] <> null))),
        EffortCheckResult("Unique preferred AIN employee Resources", () => Table.RowCount(Contracts) - Table.RowCount(Table.Distinct(Contracts, {"EmployeeID"}))),
        EffortCheckResult("Complete preferred AIN identity and contract coverage", () => Table.RowCount(Table.SelectRows(Contracts, each
            let IsCasualFallback = [#"Limit Basis"] <> null and Text.StartsWith([#"Limit Basis"], "Casual fallback", Comparer.OrdinalIgnoreCase)
            in [EmployeeID] = null or [Name] = null or [Resource] = null or [#"Facility-Abbrev"] = null or [EmploymentType] = null
                or [#"Worker Record Status"] <> "Found" or [#"Worker Contract Issue"] <> null
                or [#"Limit Basis"] = null or [#"Effective Shift Cap"] = null
                or (not IsCasualFallback and ([#"Contracted FN Hours"] = null or [#"Contracted FN Hours"] <= 0
                    or [#"Contracted Roster Hours"] = null or [#"Contracted Roster Hours"] <= 0))))),
        EffortCheckResult("Staff contracts align with Settings roster", () => Table.RowCount(Table.SelectRows(Contracts, each
            [#"Roster Start"] = null or [#"Roster End"] = null or [#"Roster Fortnights"] = null
            or [#"Roster Start"] <> RosterStart or [#"Roster End"] <> RosterEnd
            or Number.Abs([#"Roster Fortnights"] - RosterFortnights) > Tolerance))),
        EffortCheckResult("Allocated employees aggregate once", () => Table.RowCount(AllocationTotals) - Table.RowCount(Table.Distinct(AllocationTotals, {"EmployeeID"}))),
        EffortCheckResult("Exact hour and shift-equivalent calculations", () => Table.RowCount(Table.SelectRows(Prepared, each
            Number.Abs([#"Allocated Shift Equivalent"] - [#"Allocated Roster Hours"] / ShiftDuration) > Tolerance
            or ([#"Contracted Roster Hours"] <> null and Number.Abs([#"Over-Contract Hours"] - ([#"Allocated Roster Hours"] - [#"Contracted Roster Hours"])) > Tolerance)
            or ([#"Over-Contract Hours"] <> null and Number.Abs([#"Over-Contract Shift Equivalent"] - [#"Over-Contract Hours"] / ShiftDuration) > Tolerance)))),
        EffortCheckResult("Unique over-contract employee Resource roster keys", () => Table.RowCount(PublishedRows) -
            Table.RowCount(Table.Distinct(PublishedRows, {"Facility", "EmployeeID", "Resource", "Roster Start", "Roster End"}))),
        EffortCheckResult("Published rows are positive exact-hour overages", () => Table.RowCount(Table.SelectRows(PublishedRows, each
            [#"Over-Contract Hours"] = null or [#"Over-Contract Hours"] <= Tolerance or not [#"Is Over Contract"]
            or [#"Allocated Role List"] = null or Text.Trim([#"Allocated Role List"]) = ""))),
        EffortCheckResult("Published rows are Unit1 AIN only", () => Table.RowCount(Table.SelectRows(PublishedRows, each
            [#"Coverage Scope"] <> "Unit1/AIN" or [#"Preferred Role"] <> "AIN" or [Facility] <> "Unit1")))
    },
    DiagnosticChecks = {
        [Check = "Casual fallback contracts excluded from over-contract", Status = if Table.IsEmpty(CasualFallbackRows) then "Pass" else "Warning",
            Failures = Table.RowCount(CasualFallbackRows), Details = "These employees have no exact positive contract hours and are not classified as over-contract."]
    }
in
    Table.Buffer(Table.FromRecords(BlockingChecks & DiagnosticChecks,
        type table [Check = text, Status = text, Failures = nullable number, Details = nullable text]));

// Query: ResourceContractOverage
// Purpose: Publish positive exact-hours contract overages for preferred Unit1/AIN Resources.
// Output: One row per employee/preferred Resource/roster; load this query to an Excel table named ResourceContractOverage.
shared ResourceContractOverage = let
    Failures = Table.SelectRows(ResourceContractOverage_CHECK, each [Status] = "Fail"),
    Checked = if Table.IsEmpty(Failures) then ResourceContractOverage_Prepare
        else error Error.Record("Effort contract overage validation", "Required ResourceContractOverage checks failed.", Failures),
    PositiveOverages = Table.SelectRows(Checked, each [#"Is Over Contract"]),
    Output = Table.SelectColumns(PositiveOverages, {"Facility", "Facility-Abbrev", "Preferred Role", "Resource", "EmployeeID", "Name", "Employment Type",
        "Roster Start", "Roster End", "Roster Fortnights", "Contracted FN Hours", "Contracted Roster Hours", "Contracted Shift Equivalent",
        "Contracted Shifts", "Settings Max Availability", "Effective Shift Cap", "Allocated Roster Hours", "Allocated Shift Equivalent",
        "Over-Contract Hours", "Over-Contract Shift Equivalent", "Allocated Role List", "Limit Basis", "Is Over Contract", "Coverage Scope"})
in
    Table.Sort(Output, {{"Over-Contract Hours", Order.Descending}, {"Resource", Order.Ascending}});

// Query: CentriSyncPaths
// Purpose: Preserve the established mapping-table interface used by existing Excel connections.
// Notes: Compatibility reference only; the public mapping workbook is opened once by IMPORT CentriSyncPaths.
shared CentriSyncPaths = #"IMPORT CentriSyncPaths";