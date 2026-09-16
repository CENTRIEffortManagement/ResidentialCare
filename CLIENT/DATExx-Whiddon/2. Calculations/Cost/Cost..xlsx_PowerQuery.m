// Power Query from: Cost..xlsx
// Pathname: CLIENT\DATExx-Whiddon\2. Calculations\Cost\Cost..xlsx
// Extracted: 2026-05-21T00:48:03.036Z

section Section1;

// Query: IMPORT CentriSyncPaths
// Purpose: Read the machine's shared path mapping for portable workbook imports.
shared #"IMPORT CentriSyncPaths" =
let
    Navigation = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    Mapping = Navigation{[Item="CentriSyncPaths", Kind="Table"]}[Data]
in
    Table.Buffer(Mapping);

// Query: CostPathTABLE
// Purpose: Resolve this workbook's path through the standard CentriSyncPaths mapping.
// Inputs: FilePathUrl, either the standard one-row FilePath table or this workbook's single named cell.
// Output: Resolved workbook folder and source identity for relative organisation imports.
// Notes: The worksheet output is not a path input; reading it here would reuse stale query results.
shared CostPathTABLE =
let
    FilePathUrl =
    let
        Source = Excel.CurrentWorkbook(){[Name="FilePathUrl"]}[Content],
        // Excel exposes the existing single-cell FilePathUrl name as Column1.
        // Keep the canonical FilePath column when the standard input table is installed.
        PathInput = if Table.HasColumns(Source, {"FilePath"}) then Source
            else if Table.ColumnNames(Source) = {"Column1"} then
                Table.RenameColumns(Source, {{"Column1", "FilePath"}})
            else error "FilePathUrl must be a one-row FilePath table or a single named cell.",
        SelectedColumns = Table.SelectColumns(PathInput, {"FilePath"}),
        ChangedType = Table.TransformColumnTypes(SelectedColumns, {{"FilePath", type text}}),
        ReplacedValue = Table.TransformColumns(ChangedType, {{"FilePath", each if _ = null then null else Text.Replace(_, "/", "\"), type text}}),
        ValidatedTable = if Table.RowCount(ReplacedValue) = 1 then ReplacedValue else error "FilePathUrl must contain exactly one data row.",
        BufferedTable = Table.Buffer(ValidatedTable)
    in
        BufferedTable,

    RawFilePathValue = FilePathUrl{0}[FilePath],
    RawFilePath = if RawFilePathValue = null or Text.Trim(RawFilePathValue) = "" then
        error "FilePathUrl is blank. Save this workbook and recalculate its CELL filename formula."
        else Text.Trim(RawFilePathValue),
    // CELL("filename", reference) returns folder\[workbook.xlsx]sheet.
    // Remove only the workbook brackets and sheet suffix before resolving the folder.
    WorkbookPath = if Text.Contains(RawFilePath, "[") then
        Text.BeforeDelimiter(RawFilePath, "[") & Text.BetweenDelimiters(RawFilePath, "[", "]")
        else RawFilePath,
    InputFileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    ValidatedWorkbookPath = if Comparer.OrdinalIgnoreCase(InputFileName, "Cost..xlsx") = 0 then
        WorkbookPath
        else error "FilePathUrl identifies another workbook. In the named cell use =CELL(""filename"",A1), then save and recalculate Cost..xlsx.",
    CentriSyncPaths_Table = #"IMPORT CentriSyncPaths",
    CentriSyncPaths_ChangedType = Table.TransformColumnTypes(Table.SelectColumns(CentriSyncPaths_Table, {"SharepointRootUrl", "SyncedFolderRootPath"}), {{"SharepointRootUrl", type text}, {"SyncedFolderRootPath", type text}}),
    NormalizePath = (value as nullable text) as nullable text =>
        let
            TextValue = if value = null then null else Text.From(value),
            SlashNormalized = if TextValue = null then null else Text.Replace(TextValue, "/", "\"),
            Trimmed = if SlashNormalized = null then null else Text.TrimEnd(SlashNormalized, "\")
        in
            Trimmed,
    FilePath = NormalizePath(ValidatedWorkbookPath),
    CentriSyncPaths_Normalized = Table.TransformColumns(
        CentriSyncPaths_ChangedType,
        {
            {"SharepointRootUrl", each NormalizePath(_), type text},
            {"SyncedFolderRootPath", each NormalizePath(_), type text}
        }
    ),
    SharePointCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SharepointRootUrl], type text),
    SharePointDocumentsCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\Shared Documents", type text),
    LocalCandidates = Table.AddColumn(CentriSyncPaths_Normalized, "MatchRoot", each [SyncedFolderRootPath], type text),
    MatchCandidates = Table.Combine({SharePointCandidates, SharePointDocumentsCandidates, LocalCandidates}),
    MatchCandidates_WithLength = Table.AddColumn(MatchCandidates, "MatchRootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRows = Table.SelectRows(
        MatchCandidates_WithLength,
        each [MatchRoot] <> null
            and Text.Trim([MatchRoot]) <> ""
            and [SyncedFolderRootPath] <> null
            and Text.Trim([SyncedFolderRootPath]) <> ""
            and Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)
            and (Text.Length(FilePath) = [MatchRootLength] or Text.Range(FilePath, [MatchRootLength], 1) = "\")
    ),
    SortedMatches = Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}}),
    BestMatch = if Table.RowCount(SortedMatches) > 0 then SortedMatches{0} else error "FilePathUrl did not match any CentriSyncPaths root: " & FilePath,
    RelativePath = Text.Range(FilePath, BestMatch[MatchRootLength]),
    RelativePath_Trimmed = Text.TrimStart(RelativePath, "\"),
    LocalFullPath =
        if RelativePath_Trimmed = "" then
            BestMatch[SyncedFolderRootPath]
        else
            BestMatch[SyncedFolderRootPath] & "\" & RelativePath_Trimmed,
    RootPath = Text.BeforeDelimiter(LocalFullPath, "\", {0, RelativePosition.FromEnd}),
    Result = #table({"Variable Name", "Value"},
        {{"Root Path", RootPath}, {"FilePathUrl", FilePath}, {"FileName", InputFileName}})
in
    Result;

// Query: ClientPath
// Purpose: Derive the Whiddon organisation root from the resolved Cost workbook folder.
shared ClientPath =
let
    WorkbookFolder = CostPathTABLE{[#"Variable Name"="Root Path"]}[Value],
    CostSuffix = "\2. Calculations\Cost",
    ClientFolder = if Text.EndsWith(WorkbookFolder, CostSuffix, Comparer.OrdinalIgnoreCase) then
            Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(CostSuffix))
        else error "Cost workbook folder must end in \2. Calculations\Cost.",
    Result = if Text.EndsWith(ClientFolder, "\DATExx-Whiddon", Comparer.OrdinalIgnoreCase) then ClientFolder
        else error "Cost FilePathUrl must identify the DATExx-Whiddon client."
in
    Result;

// Query: IMPORT Settings Data
// Purpose: Read Whiddon Unit1 settings once for the cost duration and calendar.
// Inputs: ClientPath, resolved from this workbook's FilePathUrl through CentriSyncPaths.
// Notes: Unit1 remains the approved standard-FTE and calendar authority.
shared #"IMPORT Settings Data" =
    Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
        ClientPath & "\UNITS\Unit1\2. Calculations\Settings Data.xlsx")), null, true));

// Query: IMPORT Inefficiencies
// Purpose: Read saved organisation inefficiency results using the same resolved client root.
shared #"IMPORT Inefficiencies" =
    Table.Buffer(Excel.Workbook(Binary.Buffer(File.Contents(
        ClientPath & "\2. Calculations\E-O-I\Inefficiencies.xlsx")), null, true));

// Query: EXTRACT ShiftDuration
// Purpose: Select the standard-FTE hours table from the shared Settings import.
shared #"EXTRACT ShiftDuration" =
let
    Matches = Table.SelectRows(#"IMPORT Settings Data",
        each [Item] = "ShiftDuration" and [Kind] = "Table"),
    Result = if Table.RowCount(Matches) = 1 then Matches{0}[Data]
        else error "Settings Data must contain exactly one ShiftDuration table."
in
    Result;

// Query: EXTRACT PermutationDimensions
// Purpose: Select the planning calendar from the same Unit1 Settings import.
shared #"EXTRACT PermutationDimensions" =
let
    Source = #"IMPORT Settings Data"{[Item="PermutationDimensions", Kind="Table"]}[Data],
    TypedCalendar = Table.TransformColumnTypes(Source,
        {{"Date", type date}, {"Day", Int64.Type}, {"Shifts", type text}, {"Period", Int64.Type}, {"RolesList", type text}})
in
    TypedCalendar;

// Query: ShiftDuration
// Purpose: Validate the configured standard-FTE duration in hours.
// Inputs: EXTRACT ShiftDuration; exactly one numeric, positive, finite hours value.
// Notes: Invalid settings fail; no fixed duration fallback is applied.
shared ShiftDuration =
let
    Data = #"EXTRACT ShiftDuration",
    RequiredColumn = if Table.HasColumns(Data, {"ShiftDuration"}) then Data
        else error "The ShiftDuration table must contain the ShiftDuration column.",
    Hours = if Table.RowCount(RequiredColumn) = 1 then RequiredColumn{0}[ShiftDuration]
        else error "ShiftDuration must contain exactly one data row.",
    ValidatedHours = if not Value.Is(Hours, type number) then
            error "ShiftDuration must be a numeric hours value."
        else if Number.IsNaN(Hours) or Hours <= 0 or Number.Abs(Hours) = #infinity then
            error "ShiftDuration must be a positive finite number of hours."
        else Hours
in
    ValidatedHours;

// Query: ShiftHrs
// Purpose: Retain the existing parameter interface as an alias for ShiftDuration hours.
shared ShiftHrs = ShiftDuration meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

// Query: PermutationDimensions
// Purpose: Preserve the existing calendar interface for TargetRatebyDate.
shared PermutationDimensions = #"EXTRACT PermutationDimensions";

shared TargetRatebyDate = let
    Source = PermutationDimensions,
    #"Renamed Columns" = Table.RenameColumns(Source,{{"Date", "DateX"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "Date", each Date.AddDays([DateX] ,
28)),
    #"Grouped Rows" = Table.Group(#"Added Custom1", {"Date"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Grouped Rows",{"Count"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns", "TargetRate", each TargetDailyRate),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"Date", type date}})
in
    #"Changed Type";

shared RatesDelta = let
    Source = Excel.CurrentWorkbook(){[Name="RatesDelta"]}[Content],
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(Source, {"Role", "Period"}, "Attribute", "Value")
in
    #"Unpivoted Other Columns";

// Query: EffortOutcomesAG1_1DayShiftAB
// Purpose: Select and prepare the single required table from IMPORT Inefficiencies.
// Notes: Keep the existing output schema; one extraction needs no intermediate EXTRACT query.
shared EffortOutcomesAG1_1DayShiftAB = let
    Source = #"IMPORT Inefficiencies",
    EffortOutcomesAG1_1DayShiftAB_Table = Source{[Item="EffortOutcomesAG1_1DayShiftAB",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(EffortOutcomesAG1_1DayShiftAB_Table,{{"Level", type text}, {"Potential", Int64.Type}, {"Planned", Int64.Type}, {"Degree", Int64.Type}, {"Spare", Int64.Type}, {"Spare Allocated", Int64.Type}, {"Block", Int64.Type}, {"EXCESS OVER-ALLOCATION", type number}, {"SPARE STRETCH", type number}, {"SPARE SLACK", type number}, {"SPARE STRETCH ALLOCATED", type number}, {"SPARE SLACK ALLOCATED", type number}, {"EXCESS STRETCH (ALLOCATED)", type number}, {"EXCESS ALLOCATED SLACK", type number}, {"EXCESS STRETCH", type number}, {"EXCESS SLACK", type number}, {"POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"POTENTIAL SHORTFALL", type number}, {"WASTED STRETCH", type number}, {"WASTED SLACK", type number}, {"ALLOCATED STRETCH", type number}, {"ALLOCATED SLACK", type number}, {"UNALLOCATED SLACK", type number}, {"LATENT", Int64.Type}, {"LATENT.ALLOCATED", Int64.Type}, {"LATENT.OVERALLOCATED", Int64.Type}, {"AB Potential", Int64.Type}, {"Check", type number}, {"wAB EXCESS OVER-ALLOCATION", type number}, {"wAB SPARE STRETCH", type number}, {"wAB SPARE SLACK", type number}, {"wAB SPARE STRETCH ALLOCATED", type number}, {"wAB SPARE SLACK ALLOCATED", type number}, {"wAB EXCESS STRETCH (ALLOCATED)", type number}, {"wAB EXCESS ALLOCATED SLACK", type number}, {"wAB EXCESS STRETCH", type number}, {"wAB EXCESS SLACK", type number}, {"wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", type number}, {"wAB POTENTIAL SHORTFALL", type number}, {"wAB WASTED STRETCH", type number}, {"wAB WASTED SLACK", type number}, {"wAB ALLOCATED STRETCH", type number}, {"wAB ALLOCATED SLACK", type number}, {"wAB UNALLOCATED SLACK", type number}, {"wAB LATENT", Int64.Type}, {"wAB LATENT.ALLOCATED", Int64.Type}, {"wAB LATENT.OVERALLOCATED", Int64.Type}, {"wPotential", Int64.Type}, {"Column1", type any}, {"Column2", type any}, {"Date", type date}, {"L1.1", type text}, {"L1.2", type text}, {"D", type number}, {"Cs", type number}, {"CX", type number}, {"A", type number}, {"Apn", type number}, {"Apx", type number}, {"Ai", type number}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Potential", "Planned", "Degree", "Spare", "Spare Allocated", "Block"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Date", "L1.1", "L1.2", "D", "Cs", "CX", "A", "Apn", "Apx", "Ai", "Level", "EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "AB Potential", "Check", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential", "Column1", "Column2"}),
    #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns",{"Level", "Check", "wAB EXCESS OVER-ALLOCATION", "wAB SPARE STRETCH", "wAB SPARE SLACK", "wAB SPARE STRETCH ALLOCATED", "wAB SPARE SLACK ALLOCATED", "wAB EXCESS STRETCH (ALLOCATED)", "wAB EXCESS ALLOCATED SLACK", "wAB EXCESS STRETCH", "wAB EXCESS SLACK", "wAB POTENTIAL SHORTFALL (OVER-ALLOCATED)", "wAB POTENTIAL SHORTFALL", "wAB WASTED STRETCH", "wAB WASTED SLACK", "wAB ALLOCATED STRETCH", "wAB ALLOCATED SLACK", "wAB UNALLOCATED SLACK", "wAB LATENT", "wAB LATENT.ALLOCATED", "wAB LATENT.OVERALLOCATED", "wPotential", "Column1", "Column2"})
in
    #"Removed Columns1";

shared Rates = let
    Source = Excel.CurrentWorkbook(){[Name="Rates"]}[Content],
    #"Added Custom" = Table.AddColumn(Source, "$OT", each [OTLoading]*[Rate]),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "$Agency", each [AgencyPremium]*[Rate]),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom1",{ "OTLoading", "AgencyPremium"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Columns",{{"Rate", "$Rate"}})
in
    #"Renamed Columns";

shared TargetDailyRate = 5 meta [IsParameterQuery=true, Type="Any", IsParameterQueryRequired=true];

// Query: CostAG1_1DayShiftAB
// Purpose: Convert effort quantities to costs using configured standard-FTE hours and existing rates.
shared CostAG1_1DayShiftAB = let
    Source = EffortOutcomesAG1_1DayShiftAB,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Date] <> null)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Cs", "CX", "A", "Apn", "Apx", "Ai", "DATESHIFT", "CapacityMaxHC"}),
    EXCESSALLOCATION = Table.AddColumn(#"Removed Columns", "ExcessAllocation", each 1*(
 
[#"EXCESS OVER-ALLOCATION"]+
[SPARE STRETCH ALLOCATED]+
[SPARE SLACK ALLOCATED]+
[#"EXCESS STRETCH (ALLOCATED)"]+
[EXCESS ALLOCATED SLACK])),
    #"Added REMOVECARERforRN" = Table.AddColumn(EXCESSALLOCATION, "ExcessAllocation-Reaslistic", each if [Facility]= "Berry" and [L1.2] = "NIGHT" and [L1.1] = "Carer" then 0
else [ExcessAllocation]),
    SHORTAGE = Table.AddColumn(#"Added REMOVECARERforRN", "Shortage", each [D]*([POTENTIAL SHORTFALL]+
[WASTED STRETCH]+
[WASTED SLACK]+
[UNALLOCATED SLACK])),
    #"Added SHORTAGE-REALISTIC" = Table.AddColumn(SHORTAGE, "Shortage-Realistic", each if [Facility]= "Berry" and [L1.2] = "NIGHT" and [L1.1] = "Registered Nurse" then 0
else [Shortage]),
    PLANNEDOT = Table.AddColumn(#"Added SHORTAGE-REALISTIC", "PlannedOvertime", each [D]*[#"POTENTIAL SHORTFALL (OVER-ALLOCATED)"]),
    #"Removed Columns1" = Table.RemoveColumns(PLANNEDOT,{"EXCESS OVER-ALLOCATION", "SPARE STRETCH", "SPARE SLACK", "SPARE STRETCH ALLOCATED", "SPARE SLACK ALLOCATED", "EXCESS STRETCH (ALLOCATED)", "EXCESS ALLOCATED SLACK", "EXCESS STRETCH", "EXCESS SLACK", "POTENTIAL SHORTFALL (OVER-ALLOCATED)", "POTENTIAL SHORTFALL", "WASTED STRETCH", "WASTED SLACK", "ALLOCATED STRETCH", "ALLOCATED SLACK", "UNALLOCATED SLACK", "LATENT", "LATENT.ALLOCATED", "LATENT.OVERALLOCATED", "AB Potential"}),
    #"Inserted Day Name" = Table.AddColumn(#"Removed Columns1", "Day Name", each Date.DayOfWeekName([Date]), type text),
    #"Merged Queries" = Table.NestedJoin(#"Inserted Day Name", {"L1.1", "L1.2", "Day Name"}, RatesDelta, {"Role", "Period", "Attribute"}, "Rates", JoinKind.LeftOuter),
    #"Expanded Rates1" = Table.ExpandTableColumn(#"Merged Queries", "Rates", {"Attribute", "Value"}, {"Rates.Attribute", "Rates.Value"}),
    // ShiftDuration converts standard-FTE quantities to hours before applying the existing rates.
    UNNECESSARYALLOCATION = Table.AddColumn(#"Expanded Rates1", "UnnecessaryAllocation$", each 23
*[#"ExcessAllocation-Reaslistic"]*ShiftDuration),
    SHORTAGEREPLACE = Table.AddColumn(UNNECESSARYALLOCATION, "ShortageReplacement$", each [#"Shortage-Realistic"]*
[Rates.Value]
*ShiftDuration),
    #"OVERTIME$" = Table.AddColumn(SHORTAGEREPLACE, "Overtime$", each [PlannedOvertime]
*0
*ShiftDuration),
    #"Replaced Value" = Table.ReplaceValue(#"OVERTIME$",null,0,Replacer.ReplaceValue,{"UnnecessaryAllocation$", "ShortageReplacement$", "Overtime$"}),
    INEFFICIENCYCOST = Table.AddColumn(#"Replaced Value", "InefficiencyCost", each [#"UnnecessaryAllocation$"]
+[#"ShortageReplacement$"]
+[#"Overtime$"])
in
    INEFFICIENCYCOST;

shared CostAG1_1DayShiftABSummary = let
    Source = CostAG1_1DayShiftAB,
    #"Grouped Rows" = Table.Group(Source, {"L1.1", "Facility", "Version"}, {{"RoleCost", each List.Sum([InefficiencyCost]), type number}, {"UnnecessaryAllocation", each List.Sum([#"UnnecessaryAllocation$"]), type number}, {"Shortage Replacement", each List.Sum([#"ShortageReplacement$"]), type number}, {"OT", each List.Sum([#"Overtime$"]), type number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Version", Order.Ascending}, {"Facility", Order.Ascending}, {"L1.1", Order.Ascending}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Sorted Rows",{"Version", "Facility", "L1.1", "RoleCost", "UnnecessaryAllocation", "Shortage Replacement", "OT"})
in
    #"Reordered Columns";

shared Saving = let
    Source = CostAG1_1DayShiftAB,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Version", "L1.2", "L1.1", "Date","InefficiencyCost","Facility"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([Version] <> null)),
    #"Pivoted Column" = Table.Pivot(#"Filtered Rows", List.Distinct(#"Filtered Rows"[Version]), "Version", "InefficiencyCost", List.Sum),
    #"Inserted Subtraction" = Table.AddColumn(#"Pivoted Column", "Saving", each [Base] - [AllocationChange], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"AllocationChange", "Base"}),
    #"Grouped Rows" = Table.Group(#"Removed Columns", {"Date", "Facility"}, {{"Savings", each List.Sum([Saving]), type number}})
in
    #"Grouped Rows";

shared SavingFacilityDay = let
    Source = CostAG1_1DayShiftAB,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Version", "L1.2", "L1.1", "Date","InefficiencyCost","Facility"}),
    #"Pivoted Column" = Table.Pivot(#"Removed Other Columns", List.Distinct(#"Removed Other Columns"[Version]), "Version", "InefficiencyCost", List.Sum),
    #"Inserted Subtraction" = Table.AddColumn(#"Pivoted Column", "Saving", each [Base] - [AllocationChange], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"AllocationChange", "Base"}),
    #"Grouped Rows" = Table.Group(#"Removed Columns", {"Date", "Facility"}, {{"Savings", each List.Sum([Saving]), type number}})
in
    #"Grouped Rows";

shared CostPerDay = let
    Source = CostAG1_1DayShiftAB,
    #"Grouped Rows" = Table.Group(Source, {"Date", "Facility"}, {{"OpportunityCostDay", each List.Sum([InefficiencyCost]), type nullable number}}),
    #"Grouped Rows1" = Table.Group(#"Grouped Rows", {"Facility"}, {{"Count", each List.Sum([OpportunityCostDay]), type nullable number}})
in
    #"Grouped Rows1";
