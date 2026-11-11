section Section1;

shared #"IMPORT OvertimeRosterChecks CentriSyncPaths" = let
    Navigation = Excel.Workbook(File.Contents("C:\Users\Public\Public Scripts\CentriSyncPaths.xlsx"), null, true),
    Mapping = Navigation{[Item = "CentriSyncPaths", Kind = "Table"]}[Data]
in
    Mapping;

shared #"INPUT OvertimeRosterChecks FilePathUrl" = Excel.CurrentWorkbook(){[Name = "FilePathUrl"]}[Content];

shared OvertimeRosterChecks_Path = let
    NormalizePath = (value as any) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(Text.From(value)), "/", "\"), "\"),
    PathInput = #"INPUT OvertimeRosterChecks FilePathUrl",
    PathColumn = if Table.HasColumns(PathInput, {"FilePath"}) then PathInput[FilePath]
        else if Table.ColumnNames(PathInput) = {"Column1"} then PathInput[Column1]
        else error "FilePathUrl must have a FilePath column or be one named cell.",
    RawPath = if List.Count(PathColumn) = 1 then NormalizePath(PathColumn{0}) else error "FilePathUrl requires exactly one value.",
    NonBlankPath = if RawPath = null or RawPath = "" then error "Save and recalculate OvertimeRosterChecks.xlsx before resolving FilePathUrl." else RawPath,
    WorkbookPath = if Text.Contains(NonBlankPath, "[") then Text.BeforeDelimiter(NonBlankPath, "[") & Text.BetweenDelimiters(NonBlankPath, "[", "]") else NonBlankPath,
    FileName = Text.AfterDelimiter(WorkbookPath, "\", {0, RelativePosition.FromEnd}),
    NamedWorkbookPath = if Comparer.OrdinalIgnoreCase(FileName, "OvertimeRosterChecks.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl must identify OvertimeRosterChecks.xlsx.",
    MappingColumns = Table.SelectColumns(#"IMPORT OvertimeRosterChecks CentriSyncPaths", {"SharepointRootUrl", "SyncedFolderRootPath"}),
    NormalizedMappings = Table.TransformColumns(MappingColumns, {{"SharepointRootUrl", NormalizePath, type nullable text}, {"SyncedFolderRootPath", NormalizePath, type nullable text}}),
    BufferedMappings = Table.Buffer(NormalizedMappings),
    SharePointRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    DocumentRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\Shared Documents", type nullable text),
    LocalRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    CandidateRoots = Table.Combine({SharePointRoots, DocumentRoots, LocalRoots}),
    RootLengths = Table.AddColumn(CandidateRoots, "RootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRoots = Table.SelectRows(RootLengths, each [RootLength] > 0 and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and Text.StartsWith(NamedWorkbookPath, [MatchRoot], Comparer.OrdinalIgnoreCase)
        and (Text.Length(NamedWorkbookPath) = [RootLength] or Text.Range(NamedWorkbookPath, [RootLength], 1) = "\")),
    OrderedRoots = Table.Sort(MatchingRoots, {{"RootLength", Order.Descending}}),
    LongestRoot = if Table.IsEmpty(OrderedRoots) then error "FilePathUrl has no CentriSyncPaths mapping." else OrderedRoots{0},
    EqualRoots = Table.SelectRows(OrderedRoots, each [RootLength] = LongestRoot[RootLength]),
    Destinations = List.Distinct(EqualRoots[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    SelectedRoot = if List.Count(Destinations) = 1 then LongestRoot else error "CentriSyncPaths has conflicting longest-prefix destinations.",
    RelativePath = Text.TrimStart(Text.Range(NamedWorkbookPath, SelectedRoot[RootLength]), "\"),
    LocalWorkbookPath = SelectedRoot[SyncedFolderRootPath] & "\" & RelativePath,
    WorkbookFolder = Text.BeforeDelimiter(LocalWorkbookPath, "\", {0, RelativePosition.FromEnd}),
    NestedSuffix = "\2. Calculations\OvertimeRosterChecks",
    ClientPath = if Text.EndsWith(WorkbookFolder, NestedSuffix, Comparer.OrdinalIgnoreCase) then Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(NestedSuffix))
        else error "OvertimeRosterChecks.xlsx must be in the organisation's 2. Calculations/OvertimeRosterChecks folder.",
    ClientSegments = Text.Split(ClientPath, "\"),
    OrganisationPath = if List.PositionOf(ClientSegments, "UNITS", Occurrence.First, Comparer.OrdinalIgnoreCase) < 0 then ClientPath
        else error "OvertimeRosterChecks.xlsx is organisation-level and cannot be located beneath UNITS.",
    Result = [ClientPath = OrganisationPath, WorkbookPath = LocalWorkbookPath]
in
    Result;

shared #"IMPORT Published Roster" = (unit as text) as table =>
let
    SourcePath = OvertimeRosterChecks_Path[ClientPath] & "\UNITS\" & unit & "\1. Input\Allocation\Published Roster.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true),
    SourceObject = Navigation{[Item = "Roster Data", Kind = "Sheet"]}[Data]
in
    SourceObject;

shared #"IMPORT Workers" = (unit as text) as table =>
let
    SourcePath = OvertimeRosterChecks_Path[ClientPath] & "\UNITS\" & unit & "\1. Input\Worker's Report.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true),
    SourceObject = Navigation{[Item = "Sheet1", Kind = "Sheet"]}[Data]
in
    SourceObject;

shared #"IMPORT Settings Data" = (unit as text) as table =>
let
    SourcePath = OvertimeRosterChecks_Path[ClientPath] & "\UNITS\" & unit & "\2. Calculations\Settings Data.xlsx",
    Navigation = Excel.Workbook(File.Contents(SourcePath), null, true)
in
    Navigation;

shared #"IMPORT OvertimeRosterChecks Whiddon Lists" = SharePoint.Tables("https://centri001.sharepoint.com/sites/WhiddonCENTRI", [Implementation = "2.0", ViewMode = "All"]);

shared #"LINK OvertimeRosterChecks Roles" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Id = "6b0b0767-44f4-4bb8-b116-f1e99b3476f0"]}[Items];

shared #"LINK OvertimeRosterChecks Facilities" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Id = "1987eadb-ad2e-491a-a927-e5585667d4c5"]}[Items];

shared #"LINK OvertimeRosterThreshold" = #"IMPORT OvertimeRosterChecks Whiddon Lists"{[Title = "OvertimeRosterThreshold"]}[Items];

shared fnOvertimeRosterChecksText = (value as any) as nullable text =>
let
    TextValue = if value = null then null else Text.Trim(Text.From(value, "en-AU")),
    Result = if TextValue = "" then null else TextValue
in
    Result;

shared fnOvertimeRosterChecksNumber = (value as any) as nullable number =>
let
    Parsed = try Number.From(value) otherwise null,
    Result = if Parsed = null then null else if Number.IsNaN(Parsed) or Number.Abs(Parsed) = #infinity then null else Parsed
in
    Result;

shared fnOvertimeRosterChecksLogical = (value as any) as nullable logical =>
let
    TextValue = fnOvertimeRosterChecksText(value),
    Key = if TextValue = null then null else Text.Upper(TextValue),
    Result = if Value.Is(value, type logical) then value else if List.Contains({"TRUE", "YES", "1"}, Key) then true
        else if List.Contains({"FALSE", "NO", "0"}, Key) then false else null
in
    Result;

shared fnOvertimeRosterChecksKey = (values as list) as text => Binary.ToText(Json.FromValue(values), BinaryEncoding.Base64);

shared fnOvertimeRosterChecksClockSafe = (start as nullable datetime, finish as nullable datetime) as logical =>
let
    ValidBounds = start <> null and finish <> null and finish >= start,
    Years = if not ValidBounds then {} else {Date.Year(Date.From(start))..Date.Year(Date.From(finish))},
    WindowsByYear = List.Transform(Years, (year) => List.Transform({4, 10}, (month) =>
        let FirstDate = #date(year, month, 1), Sunday = Date.AddDays(FirstDate, Number.Mod(7 - Date.DayOfWeek(FirstDate, Day.Sunday), 7))
        in [WindowStart = Sunday & #time(2, 0, 0), WindowEnd = Sunday & #time(3, 0, 0)])),
    Windows = List.Combine(WindowsByYear),
    CrossesClockChange = if not ValidBounds then true else List.AnyTrue(List.Transform(Windows, each start < [WindowEnd] and finish > [WindowStart]))
in
    ValidBounds and not CrossesClockChange;

shared fnOvertimeRosterChecksSheetPrepare = (source as table, unit as text, required as list) as table =>
let
    Headers = if Table.HasColumns(source, required) then source else Table.PromoteHeaders(source, [PromoteAllScalars = true]),
    MissingColumns = List.Difference(required, Table.ColumnNames(Headers)),
    RequiredSchema = if List.IsEmpty(MissingColumns) then Headers else error Error.Record("OvertimeRosterChecks source schema", "Required source columns are missing.", [SourceUnit = unit, MissingColumns = MissingColumns]),
    SourceRows = Table.AddIndexColumn(RequiredSchema, "SourceRow", 1, 1, Int64.Type),
    SourceUnit = Table.AddColumn(SourceRows, "SourceUnit", each unit, type text)
in
    SourceUnit;

shared OvertimeRosterChecks_DailyHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "DailyHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one DailyHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "DailyHoursLimit must be a positive number." else if ParsedValue <= 0 then error "DailyHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_CSHospitalityHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "CSHospitalityHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one CSHospitalityHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "CSHospitalityHoursLimit must be a positive number." else if ParsedValue <= 0 then error "CSHospitalityHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_CSPersonalCareHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "CSPersonalCareHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one CSPersonalCareHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "CSPersonalCareHoursLimit must be a positive number." else if ParsedValue <= 0 then error "CSPersonalCareHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_MinimumBreakHours = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "MinimumBreakHours"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one MinimumBreakHours row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "MinimumBreakHours must be a positive number." else if ParsedValue <= 0 then error "MinimumBreakHours has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_ShortGapReviewHours = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "ShortGapReviewHours"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one ShortGapReviewHours row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "ShortGapReviewHours must be a positive number." else if ParsedValue <= 0 then error "ShortGapReviewHours has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_MinimumFortnightDutyFreeDays = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "MinimumFortnightDutyFreeDays"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one MinimumFortnightDutyFreeDays row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "MinimumFortnightDutyFreeDays must be a positive number." else if ParsedValue <= 0 or Number.RoundDown(ParsedValue) <> ParsedValue then error "MinimumFortnightDutyFreeDays has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_FortnightHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "FortnightHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one FortnightHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "FortnightHoursLimit must be a positive number." else if ParsedValue <= 0 then error "FortnightHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_FortnightLengthDays = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "FortnightLengthDays"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one FortnightLengthDays row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "FortnightLengthDays must be a positive number." else if ParsedValue <= 0 or Number.RoundDown(ParsedValue) <> ParsedValue then error "FortnightLengthDays has an invalid numeric value." else ParsedValue
in
    ValidatedValue;

shared OvertimeRosterChecks_FortnightStartDate = let
    BDLandedSettings = #"IMPORT Settings Data"("BD"),
    TELandedSettings = #"IMPORT Settings Data"("TE"),
    JHRYLandedSettings = #"IMPORT Settings Data"("JH-RY"),
    PreparePermutationDates = (settings as table, unit as text) as table =>
        let MatchingTables = Table.SelectRows(settings, each [Item] = "PermutationDimensions" and [Kind] = "Table"),
            PermutationTable = if Table.RowCount(MatchingTables) = 1 then MatchingTables{0}[Data]
                else error Error.Record("OvertimeRosterChecks settings dates", "Expected exactly one PermutationDimensions table.", [SourceUnit = unit]),
            DateColumn = if Table.HasColumns(PermutationTable, {"Date"}) then Table.SelectColumns(PermutationTable, {"Date"})
                else error Error.Record("OvertimeRosterChecks settings dates", "PermutationDimensions must contain Date.", [SourceUnit = unit]),
            SuppliedDates = Table.SelectRows(DateColumn, each [Date] <> null),
            ValidatedDates = Table.TransformColumns(SuppliedDates, {{"Date", each try Date.From(_) otherwise
                error Error.Record("OvertimeRosterChecks settings dates", "PermutationDimensions Date contains an invalid date.", [SourceUnit = unit]), type date}})
        in ValidatedDates,
    BDDates = PreparePermutationDates(BDLandedSettings, "BD"),
    TEDates = PreparePermutationDates(TELandedSettings, "TE"),
    JHRYDates = PreparePermutationDates(JHRYLandedSettings, "JH-RY"),
    CombinedDates = Table.Combine({BDDates, TEDates, JHRYDates}),
    DateValues = CombinedDates[Date],
    FirstDate = if List.IsEmpty(DateValues) then null else List.Min(DateValues)
in
    FirstDate;

shared OvertimeRosterChecks_RosterPrepare = let
    RequiredColumns = {"Employee Code", "Employee_Code", "Employee Roster Name", "Role", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length"},
    BDLanded = #"IMPORT Published Roster"("BD"),
    TELanded = #"IMPORT Published Roster"("TE"),
    JHRYLanded = #"IMPORT Published Roster"("JH-RY"),
    BDRows = fnOvertimeRosterChecksSheetPrepare(BDLanded, "BD", RequiredColumns),
    TERows = fnOvertimeRosterChecksSheetPrepare(TELanded, "TE", RequiredColumns),
    JHRYRows = fnOvertimeRosterChecksSheetPrepare(JHRYLanded, "JH-RY", RequiredColumns),
    CombinedRows = Table.Combine({BDRows, TERows, JHRYRows}),
    NeededColumns = Table.SelectColumns(CombinedRows, {"SourceUnit", "SourceRow", "Employee Code", "Employee_Code", "Pay Company", "Employee Roster Name", "Role", "Location", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length", "Shift Type", "Employment Type", "Non Attended", "Status", "Published", "MealAvailabilityRequired", "MealStart", "MealEnd", "CoverageStart", "CoverageEnd", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}, MissingField.UseNull),
    SourceFile = Table.AddColumn(NeededColumns, "SourceFile", each "UNITS/" & [SourceUnit] & "/1. Input/Allocation/Published Roster.xlsx", type text),
    RenamedSourceColumns = Table.RenameColumns(SourceFile, {{"Employee_Code", "AllocationCode"}, {"Employee Code", "WorkerEmployeeID"}, {"Pay Company", "EmployerKey"}, {"Employee Roster Name", "SourceName"}, {"Role", "SourceRole"}, {"Date", "SourceDate"}, {"Shift Type", "SourceShiftType"}, {"Employment Type", "SourceEmploymentType"}, {"Break Length Minutes", "BreakMinutes"}, {"Shift Net Length", "NetRosterHours"}}),
    NormalizedText = Table.TransformColumns(RenamedSourceColumns, {{"AllocationCode", fnOvertimeRosterChecksText, type nullable text}, {"WorkerEmployeeID", fnOvertimeRosterChecksText, type nullable text}, {"EmployerKey", fnOvertimeRosterChecksText, type nullable text}, {"SourceName", fnOvertimeRosterChecksText, type nullable text}, {"SourceRole", fnOvertimeRosterChecksText, type nullable text}, {"Location", fnOvertimeRosterChecksText, type nullable text}, {"SourceEmploymentType", fnOvertimeRosterChecksText, type nullable text}, {"AgreementStream", fnOvertimeRosterChecksText, type nullable text}, {"Grade", fnOvertimeRosterChecksText, type nullable text}}),
    ParsedQuantities = Table.TransformColumns(NormalizedText, {{"BreakMinutes", fnOvertimeRosterChecksNumber, type nullable number}, {"NetRosterHours", fnOvertimeRosterChecksNumber, type nullable number}, {"SourceDate", each try Date.From(_) otherwise null, type nullable date}, {"Start Time", each try Time.From(_) otherwise null, type nullable time}, {"End Time", each try Time.From(_) otherwise null, type nullable time}}),
    ParsedFacts = Table.TransformColumns(ParsedQuantities, {{"MealAvailabilityRequired", fnOvertimeRosterChecksLogical, type nullable logical}, {"OvertimeExempt", fnOvertimeRosterChecksLogical, type nullable logical}, {"BrokenShiftException", fnOvertimeRosterChecksLogical, type nullable logical}, {"Published", fnOvertimeRosterChecksLogical, type nullable logical}, {"Non Attended", fnOvertimeRosterChecksLogical, type nullable logical}}),
    PublishedDutyCandidates = Table.SelectRows(ParsedFacts, each [Published] <> false and [Non Attended] <> true),
    ShiftStart = Table.AddColumn(PublishedDutyCandidates, "ShiftStart", each if [SourceDate] = null or [Start Time] = null then null else [SourceDate] & [Start Time], type nullable datetime),
    ShiftFinish = Table.AddColumn(ShiftStart, "ShiftFinish", each if [SourceDate] = null or [End Time] = null or [Start Time] = null then null
        else (if [End Time] < [Start Time] then Date.AddDays([SourceDate], 1) else [SourceDate]) & [End Time], type nullable datetime),
    GrossHours = Table.AddColumn(ShiftFinish, "GrossHours", each if [ShiftStart] = null or [ShiftFinish] = null then null else Duration.TotalHours([ShiftFinish] - [ShiftStart]), type nullable number),
    MealStartDateTime = Table.AddColumn(GrossHours, "MealStartDateTime", each
        let Raw = [MealStart], Parsed = if Value.Is(Raw, type time) and [SourceDate] <> null then [SourceDate] & Raw else try DateTime.From(Raw) otherwise null,
            Aligned = if Parsed <> null and [ShiftStart] <> null and Parsed < [ShiftStart] and Value.Is(Raw, type time) then Parsed + #duration(1, 0, 0, 0) else Parsed
        in Aligned, type nullable datetime),
    MealFinishDateTime = Table.AddColumn(MealStartDateTime, "MealFinishDateTime", each
        let Raw = [MealEnd], Parsed = if Value.Is(Raw, type time) and [SourceDate] <> null then [SourceDate] & Raw else try DateTime.From(Raw) otherwise null,
            Aligned = if Parsed <> null and [MealStartDateTime] <> null and Parsed <= [MealStartDateTime] and Value.Is(Raw, type time) then Parsed + #duration(1, 0, 0, 0) else Parsed
        in Aligned, type nullable datetime),
    DeclaredCoverage = Table.TransformColumns(MealFinishDateTime, {{"CoverageStart", each try DateTime.From(_) otherwise null, type nullable datetime},
        {"CoverageEnd", each if Value.Is(_, type date) then DateTime.From(Date.AddDays(_, 1)) else try DateTime.From(_) otherwise null, type nullable datetime}}),
    SourceRecordID = Table.AddColumn(DeclaredCoverage, "SourceRecordID", each fnOvertimeRosterChecksKey({[SourceUnit], [SourceFile], [SourceRow]}), type text),
    NarrowedSource = Table.RemoveColumns(SourceRecordID, {"Start Time", "End Time", "MealStart", "MealEnd"})
in
    NarrowedSource;

shared OvertimeRosterChecks_WorkerPrepare = let
    RequiredColumns = {"Employee Code"},
    BDLanded = #"IMPORT Workers"("BD"),
    TELanded = #"IMPORT Workers"("TE"),
    JHRYLanded = #"IMPORT Workers"("JH-RY"),
    BDRows = fnOvertimeRosterChecksSheetPrepare(BDLanded, "BD", RequiredColumns),
    TERows = fnOvertimeRosterChecksSheetPrepare(TELanded, "TE", RequiredColumns),
    JHRYRows = fnOvertimeRosterChecksSheetPrepare(JHRYLanded, "JH-RY", RequiredColumns),
    CombinedRows = Table.Combine({BDRows, TERows, JHRYRows}),
    SelectedFacts = Table.SelectColumns(CombinedRows, {"Employee Code", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Contracted FN Hours", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}, MissingField.UseNull),
    NamedFacts = Table.RenameColumns(SelectedFacts, {{"Employee Code", "WorkerEmployeeID"}, {"Position Description", "WorkerPosition"}, {"Employee Status Description", "WorkerEmploymentType"}, {"AgreementStream", "WorkerAgreementStream"}, {"Grade", "WorkerGrade"}, {"OvertimeExempt", "WorkerOvertimeExempt"}, {"BrokenShiftException", "WorkerBrokenShiftException"}}),
    NormalizedText = Table.TransformColumns(NamedFacts, {{"WorkerEmployeeID", fnOvertimeRosterChecksText, type nullable text}, {"WorkerPosition", fnOvertimeRosterChecksText, type nullable text}, {"WorkerEmploymentType", fnOvertimeRosterChecksText, type nullable text}, {"WorkerAgreementStream", fnOvertimeRosterChecksText, type nullable text}, {"WorkerGrade", fnOvertimeRosterChecksText, type nullable text}}),
    ParsedFacts = Table.TransformColumns(NormalizedText, {{"Start Date", each try Date.From(_) otherwise null, type nullable date}, {"Termination Date", each try Date.From(_) otherwise null, type nullable date}, {"Contracted FN Hours", fnOvertimeRosterChecksNumber, type nullable number}, {"WorkerOvertimeExempt", fnOvertimeRosterChecksLogical, type nullable logical}, {"WorkerBrokenShiftException", fnOvertimeRosterChecksLogical, type nullable logical}}),
    IdentifiedWorkers = Table.SelectRows(ParsedFacts, each [WorkerEmployeeID] <> null),
    DistinctFacts = Table.Distinct(IdentifiedWorkers),
    FactFields = {"WorkerPosition", "WorkerEmploymentType", "Start Date", "Termination Date", "Contracted FN Hours", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException"},
    EmptyFacts = Record.FromList(List.Repeat({null}, List.Count(FactFields)), FactFields),
    GroupedWorkers = Table.Group(DistinctFacts, {"WorkerEmployeeID"}, {{"WorkerMatchCount", each Table.RowCount(_), Int64.Type}, {"WorkerFacts", each if Table.RowCount(_) = 1 then _{0} else EmptyFacts, type record}}),
    ExpandedFacts = Table.ExpandRecordColumn(GroupedWorkers, "WorkerFacts", {"WorkerPosition", "WorkerEmploymentType", "Start Date", "Termination Date", "Contracted FN Hours", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException"}),
    BufferedWorkers = Table.Buffer(ExpandedFacts)
in
    BufferedWorkers;

shared OvertimeRosterChecks_ShiftPrepare = let
    NormalizeKey = (value as any) as nullable text =>
        let TextValue = fnOvertimeRosterChecksText(value)
        in if TextValue = null then null else Text.Remove(Text.Upper(TextValue), {" ", "-", "_"}),
    PreparePeriods = (source as table, unit as text) as table =>
        let Selected = Table.SelectColumns(source, {"ShiftPeriod", "StartDay"}),
            NamedUnit = Table.AddColumn(Selected, "SourceUnit", each unit, type text),
            PeriodKeys = Table.TransformColumns(NamedUnit, {{"ShiftPeriod", each if NormalizeKey(_) = "NIGHT" then "NS" else NormalizeKey(_), type nullable text}, {"StartDay", fnOvertimeRosterChecksNumber, type nullable number}})
        in PeriodKeys,
    BDLandedSettings = #"IMPORT Settings Data"("BD"),
    TELandedSettings = #"IMPORT Settings Data"("TE"),
    JHRYLandedSettings = #"IMPORT Settings Data"("JH-RY"),
    BDLandedPeriods = BDLandedSettings{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    TELandedPeriods = TELandedSettings{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    JHRYLandedPeriods = JHRYLandedSettings{[Item = "ShiftPeriod", Kind = "Table"]}[Data],
    BDPeriods = PreparePeriods(BDLandedPeriods, "BD"),
    TEPeriods = PreparePeriods(TELandedPeriods, "TE"),
    JHRYPeriods = PreparePeriods(JHRYLandedPeriods, "JH-RY"),
    CombinedPeriods = Table.Combine({BDPeriods, TEPeriods, JHRYPeriods}),
    Periods = Table.Buffer(CombinedPeriods),
    RoleColumns = Table.SelectColumns(#"LINK OvertimeRosterChecks Roles", {"Roster Roles", "Role Group"}),
    RoleKeyRows = Table.AddColumn(RoleColumns, "RosterRoleKey", each NormalizeKey([Roster Roles]), type nullable text),
    // These small lookup snapshots contain only the mapping fields used by the joins.
    RoleKeys = Table.Buffer(RoleKeyRows),
    FacilitySource = #"LINK OvertimeRosterChecks Facilities",
    // SharePoint exposes the facility-name column by its current display name, Location.
    FacilityColumns = Table.SelectColumns(FacilitySource, {"Location", if Table.HasColumns(FacilitySource, {"Facility-Abbrev"}) then "Facility-Abbrev" else "field_1"}),
    FacilityNames = if Table.HasColumns(FacilityColumns, {"Facility-Abbrev"}) then FacilityColumns else Table.RenameColumns(FacilityColumns, {{"field_1", "Facility-Abbrev"}}),
    FacilityKeyRows = Table.AddColumn(FacilityNames, "LocationKey", each NormalizeKey([Location]), type nullable text),
    FacilityKeys = Table.Buffer(FacilityKeyRows),
    // Reuse the already narrowed roster across coverage, aliases and the main shift pipeline.
    SourceRows = Table.Buffer(OvertimeRosterChecks_RosterPrepare),
    UnitCoverage = Table.Group(SourceRows, {"SourceUnit"}, {
        {"Starts", each List.Distinct(List.RemoveNulls([CoverageStart])), type list},
        {"Ends", each List.Distinct(List.RemoveNulls([CoverageEnd])), type list}}),
    DeclaredUnits = Table.SelectRows(UnitCoverage, each List.Count([Starts]) = 1 and List.Count([Ends]) = 1 and [Starts]{0} < [Ends]{0}),
    AllUnitsDeclared = List.Sort(DeclaredUnits[SourceUnit]) = List.Sort({"BD", "TE", "JH-RY"}),
    CoverageStartAll = if AllUnitsDeclared then List.Max(List.Transform(DeclaredUnits[Starts], each _{0})) else null,
    CoverageEndAll = if AllUnitsDeclared then List.Min(List.Transform(DeclaredUnits[Ends], each _{0})) else null,
    SourceRoleKeys = Table.AddColumn(SourceRows, "RosterRoleKey", each NormalizeKey([SourceRole]), type nullable text),
    RoleJoin = Table.NestedJoin(SourceRoleKeys, {"RosterRoleKey"}, RoleKeys, {"RosterRoleKey"}, "RoleMatches", JoinKind.LeftOuter),
    MappedRole = Table.AddColumn(RoleJoin, "Role", each if Table.RowCount([RoleMatches]) = 1 then NormalizeKey([RoleMatches]{0}[Role Group]) else null, type nullable text),
    LocationKeys = Table.AddColumn(MappedRole, "LocationKey", each NormalizeKey([Location]), type nullable text),
    FacilityJoin = Table.NestedJoin(LocationKeys, {"LocationKey"}, FacilityKeys, {"LocationKey"}, "FacilityMatches", JoinKind.LeftOuter),
    Facility = Table.AddColumn(FacilityJoin, "Facility", each if Table.RowCount([FacilityMatches]) = 1 then fnOvertimeRosterChecksText([FacilityMatches]{0}[#"Facility-Abbrev"]) else null, type nullable text),
    ExportUnit = Table.RenameColumns(Facility, {{"SourceUnit", "SourceExportUnit"}}),
    PhysicalUnit = Table.AddColumn(ExportUnit, "SourceUnit", each [Facility], type nullable text),
    WorkerJoin = Table.NestedJoin(PhysicalUnit, {"WorkerEmployeeID"}, OvertimeRosterChecks_WorkerPrepare, {"WorkerEmployeeID"}, "WorkerMatches", JoinKind.LeftOuter),
    ExpandedWorker = Table.ExpandTableColumn(WorkerJoin, "WorkerMatches", {"WorkerMatchCount", "WorkerEmploymentType", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException", "Start Date", "Termination Date"}),
    EmployeeID = Table.AddColumn(ExpandedWorker, "EmployeeID", each if [WorkerMatchCount] = 1 then [WorkerEmployeeID] else null, type nullable text),
    AliasPairs = Table.SelectColumns(SourceRows, {"EmployerKey", "AllocationCode", "WorkerEmployeeID"}),
    DistinctAliasPairs = Table.Distinct(AliasPairs),
    AliasCounts = Table.Group(DistinctAliasPairs, {"EmployerKey", "AllocationCode"}, {{"AliasWorkerCount", each List.Count(List.Distinct(List.RemoveNulls([WorkerEmployeeID]))), Int64.Type}}),
    AliasJoin = Table.NestedJoin(EmployeeID, {"EmployerKey", "AllocationCode"}, AliasCounts, {"EmployerKey", "AllocationCode"}, "AliasCounts", JoinKind.LeftOuter),
    AliasCount = Table.ExpandTableColumn(AliasJoin, "AliasCounts", {"AliasWorkerCount"}),
    IdentityStatus = Table.AddColumn(AliasCount, "IdentityStatus", each if [EmployeeID] <> null and [AllocationCode] <> null and [EmployerKey] <> null and [Facility] <> null and [AliasWorkerCount] = 1 then "Verified" else "Unknown", type text),
    SourceShiftID = Table.AddColumn(IdentityStatus, "SourceShiftID", each fnOvertimeRosterChecksKey({if [EmployerKey] = null then null else Text.Upper([EmployerKey]), if [EmployeeID] <> null then [EmployeeID] else [WorkerEmployeeID], [Facility], [ShiftStart], [ShiftFinish], NormalizeKey([SourceRole])}), type text),
    EmployeeKey = Table.AddColumn(SourceShiftID, "EmployeeKey", each if [IdentityStatus] = "Verified" then fnOvertimeRosterChecksKey({Text.Upper([EmployerKey]), [EmployeeID]}) else "Unknown:" & [SourceRecordID], type text),
    EmploymentType = Table.AddColumn(EmployeeKey, "EmploymentType", each if [WorkerEmploymentType] <> null then [WorkerEmploymentType] else [SourceEmploymentType], type nullable text),
    ResolvedStream = Table.AddColumn(EmploymentType, "ResolvedAgreementStream", each if [AgreementStream] <> null then [AgreementStream] else [WorkerAgreementStream], type nullable text),
    ResolvedGrade = Table.AddColumn(ResolvedStream, "ResolvedGrade", each if [Grade] <> null then [Grade] else [WorkerGrade], type nullable text),
    ResolvedExempt = Table.AddColumn(ResolvedGrade, "ResolvedOvertimeExempt", each if [OvertimeExempt] <> null then [OvertimeExempt] else [WorkerOvertimeExempt], type nullable logical),
    ResolvedBroken = Table.AddColumn(ResolvedExempt, "ResolvedBrokenShiftException", each if [BrokenShiftException] <> null then [BrokenShiftException] else [WorkerBrokenShiftException], type nullable logical),
    RemovedSourcePolicies = Table.RemoveColumns(ResolvedBroken, {"AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}),
    NamedPolicies = Table.RenameColumns(RemovedSourcePolicies, {{"ResolvedAgreementStream", "AgreementStream"}, {"ResolvedGrade", "Grade"}, {"ResolvedOvertimeExempt", "OvertimeExempt"}, {"ResolvedBrokenShiftException", "BrokenShiftException"}}),
    PolicyStatus = Table.AddColumn(NamedPolicies, "PolicyStatus", each
        if ([AgreementStream] <> null and [WorkerAgreementStream] <> null and NormalizeKey([AgreementStream]) <> NormalizeKey([WorkerAgreementStream]))
            or ([Grade] <> null and [WorkerGrade] <> null and NormalizeKey([Grade]) <> NormalizeKey([WorkerGrade]))
            or ([OvertimeExempt] <> null and [WorkerOvertimeExempt] <> null and [OvertimeExempt] <> [WorkerOvertimeExempt])
            or ([BrokenShiftException] <> null and [WorkerBrokenShiftException] <> null and [BrokenShiftException] <> [WorkerBrokenShiftException]) then "Unknown" else "Known", type text),
    PublicationStatus = Table.AddColumn(PolicyStatus, "PublicationStatus", each if [Published] = true then "Known" else "Unknown", type text),
    DutyStatus = Table.AddColumn(PublicationStatus, "DutyStatus", each
        let StatusValue = NormalizeKey([Status]), Cancelled = StatusValue <> null and (Text.Contains(StatusValue, "CANCEL") or Text.Contains(StatusValue, "DRAFT")),
            ValidSpan = [ShiftStart] <> null and [ShiftFinish] <> null and [ShiftFinish] > [ShiftStart]
        in if [PublicationStatus] = "Known" and [Non Attended] = false and not Cancelled and ValidSpan then "Known" else "Unknown", type text),
    RosterScopeStatus = Table.AddColumn(DutyStatus, "RosterScopeStatus", each if [Role] <> null and not List.Contains({"AIN", "EN", "RN", "CS"}, [Role]) then "Outside scope"
        else if [IdentityStatus] = "Verified" and [DutyStatus] = "Known" and List.Contains({"AIN", "EN", "RN", "CS"}, [Role])
            and List.Contains({"PERMANENTFULLTIME", "PERMANENTPARTTIME", "FULLTIME", "PARTTIME", "FIXEDTERM", "CASUAL"}, NormalizeKey([EmploymentType])) then "Known" else "Unknown", type text),
    EligibilityStatus = Table.AddColumn(RosterScopeStatus, "EligibilityStatus", each
        if [Role] <> null and not List.Contains({"AIN", "EN", "RN", "CS"}, [Role]) then "Outside scope"
        else if [PolicyStatus] <> "Known" then "Unknown"
        else if [OvertimeExempt] = true then "Exempt"
        else if [IdentityStatus] <> "Verified" or [SourceDate] = null or [Role] = null or [Grade] = null or [OvertimeExempt] = null
            or not List.Contains({"PERSONALCARE", "CLINICAL", "HOSPITALITY"}, NormalizeKey([AgreementStream]))
            or (List.Contains({"AIN", "EN", "RN"}, [Role]) and not List.Contains({"PERSONALCARE", "CLINICAL"}, NormalizeKey([AgreementStream])))
            or ([Role] = "CS" and not List.Contains({"PERSONALCARE", "HOSPITALITY"}, NormalizeKey([AgreementStream])))
            or not List.Contains({"PERMANENTFULLTIME", "PERMANENTPARTTIME", "FULLTIME", "PARTTIME", "FIXEDTERM", "CASUAL"}, NormalizeKey([EmploymentType])) then "Unknown"
        else if ([Start Date] <> null and [SourceDate] < [Start Date]) or ([Termination Date] <> null and [SourceDate] > [Termination Date]) then "Unknown"
        else "Eligible", type text),
    DailyHoursLimit = Table.AddColumn(EligibilityStatus, "DailyHoursLimit", each if List.Contains({"AIN", "EN", "RN"}, [Role]) then OvertimeRosterChecks_DailyHoursLimit
        else if [Role] = "CS" and NormalizeKey([AgreementStream]) = "PERSONALCARE" then OvertimeRosterChecks_CSPersonalCareHoursLimit
        else if [Role] = "CS" and NormalizeKey([AgreementStream]) = "HOSPITALITY" then OvertimeRosterChecks_CSHospitalityHoursLimit else null, type nullable number),
    // Discard nested lookup tables and resolved policy helpers before grouping physical shifts.
    RemovedNestedInputs = Table.RemoveColumns(DailyHoursLimit, {"RoleMatches", "FacilityMatches", "RosterRoleKey", "LocationKey", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException", "WorkerEmploymentType"}),
    // Source-row and filename changes do not change duty identity. Identical exported copies count once.
    PhysicalFactColumns = {"WorkerEmployeeID", "AllocationCode", "EmployerKey", "SourceName", "SourceRole", "Location", "ShiftStart", "ShiftFinish", "BreakMinutes", "NetRosterHours", "SourceShiftType", "SourceEmploymentType", "Non Attended", "Status", "Published", "MealAvailabilityRequired", "MealStartDateTime", "MealFinishDateTime", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    PhysicalGroups = Table.Group(RemovedNestedInputs, {"SourceShiftID"}, {
        {"SourceRecordCount", each Table.RowCount(_), Int64.Type},
        {"SourceFactCount", each let FactColumns = Table.SelectColumns(_, PhysicalFactColumns), DistinctSourceFacts = Table.Distinct(FactColumns) in Table.RowCount(DistinctSourceFacts), Int64.Type},
        {"SourceUnits", each Text.Combine(List.Sort(List.Distinct([SourceExportUnit])), "|"), type text},
        {"SelectedSource", each let OrderedSources = Table.Sort(_, {{"SourceExportUnit", Order.Ascending}, {"SourceRow", Order.Ascending}}) in Table.First(OrderedSources), type record}}),
    PhysicalColumns = List.RemoveItems(Table.ColumnNames(RemovedNestedInputs), {"SourceShiftID"}),
    PhysicalShifts = Table.ExpandRecordColumn(PhysicalGroups, "SelectedSource", PhysicalColumns),
    ShiftMatchKey = Table.AddColumn(PhysicalShifts, "ShiftMatchKey", each fnOvertimeRosterChecksKey({[SourceUnit], [AllocationCode], [Role], [ShiftStart], [ShiftFinish]}), type text),
    // Materialise scalar physical-shift facts once before the count-and-join branch.
    BufferedMatchShifts = Table.Buffer(ShiftMatchKey),
    MatchCounts = Table.Group(BufferedMatchShifts, {"ShiftMatchKey"}, {{"MatchKeyCount", each Table.RowCount(_), Int64.Type}}),
    MatchJoin = Table.NestedJoin(BufferedMatchShifts, {"ShiftMatchKey"}, MatchCounts, {"ShiftMatchKey"}, "MatchCount", JoinKind.LeftOuter),
    MatchCount = Table.ExpandTableColumn(MatchJoin, "MatchCount", {"MatchKeyCount"}),
    MatchStatus = Table.AddColumn(MatchCount, "MatchStatus", each if [MatchKeyCount] = 1 and [SourceFactCount] = 1 and [IdentityStatus] = "Verified" and [AllocationCode] <> null and [Role] <> null and [ShiftStart] <> null and [ShiftFinish] <> null then "Unique" else "Unknown", type text),
    NetHoursStatus = Table.AddColumn(MatchStatus, "NetHoursStatus", each if [SourceFactCount] <> 1 or [NetRosterHours] = null or [NetRosterHours] < 0 or [GrossHours] = null or [GrossHours] <= 0
        or [BreakMinutes] = null or [BreakMinutes] < 0 or [NetRosterHours] > [GrossHours] + 0.0000001 then "Unknown" else "Known", type text),
    MealTreatmentStatus = Table.AddColumn(NetHoursStatus, "MealTreatmentStatus", each if [NetHoursStatus] <> "Known" then "Unknown"
        else if Number.Abs([GrossHours] - [NetRosterHours]) < 0.0000001 then "Known"
        else if [MealAvailabilityRequired] <> null and Number.Abs([GrossHours] - [NetRosterHours] - [BreakMinutes] / 60) < 0.0000001 then "Known" else "Unknown", type text),
    QualifyingNetHours = Table.AddColumn(MealTreatmentStatus, "QualifyingNetHours", each if [NetHoursStatus] <> "Known" then null
        else if [MealAvailabilityRequired] = true and Number.Abs([GrossHours] - [NetRosterHours] - [BreakMinutes] / 60) < 0.0000001 then [GrossHours] else [NetRosterHours], type nullable number),
    RosterShiftPeriod = Table.AddColumn(QualifyingNetHours, "RosterShiftPeriod", each
        let Unit = [SourceUnit], UnitPeriods = Table.SelectRows(Periods, each [SourceUnit] = Unit),
            AM = Table.SelectRows(UnitPeriods, each [ShiftPeriod] = "AM"), PM = Table.SelectRows(UnitPeriods, each [ShiftPeriod] = "PM"), NS = Table.SelectRows(UnitPeriods, each [ShiftPeriod] = "NS"),
            Valid = Table.RowCount(AM) = 1 and Table.RowCount(PM) = 1 and Table.RowCount(NS) = 1 and [ShiftStart] <> null,
            StartFraction = if [ShiftStart] = null then null else Number.From(Time.From([ShiftStart]))
        in if not Valid then null else if StartFraction >= AM{0}[StartDay] and StartFraction < PM{0}[StartDay] then "AM"
            else if StartFraction >= PM{0}[StartDay] and StartFraction < NS{0}[StartDay] then "PM" else "NS", type nullable text),
    CoverageBounds = Table.AddColumn(RosterShiftPeriod, "ExportCoverageStart", each CoverageStartAll, type nullable datetime),
    CoverageFinish = Table.AddColumn(CoverageBounds, "ExportCoverageEnd", each CoverageEndAll, type nullable datetime),
    CoverageStatus = Table.AddColumn(CoverageFinish, "CoverageStatus", each if CoverageStartAll <> null and CoverageEndAll <> null and [ShiftStart] <> null and [ShiftFinish] <> null
        and [ShiftStart] >= CoverageStartAll and [ShiftFinish] <= CoverageEndAll then "Declared" else "Unknown", type text),
    PrepareChronology = (rows as table) as table =>
        let Ordered = Table.Sort(rows, {{"ShiftStart", Order.Ascending}, {"ShiftFinish", Order.Ascending}, {"SourceShiftID", Order.Ascending}}),
            Indexed = Table.AddIndexColumn(Ordered, "EmployeeShiftIndex", 0, 1, Int64.Type),
            ShiftIDs = List.Buffer(Indexed[SourceShiftID]), Finishes = List.Buffer(Indexed[ShiftFinish]), Starts = List.Buffer(Indexed[ShiftStart]), DutyValues = List.Buffer(Indexed[DutyStatus]),
            PriorMaxFinishes = List.Buffer(List.Generate(() => [Index = 0, MaxFinish = null], each [Index] < List.Count(Finishes),
                each let CurrentFinish = Finishes{[Index]}, NextMax = if [MaxFinish] = null then CurrentFinish else if CurrentFinish = null then [MaxFinish] else List.Max({[MaxFinish], CurrentFinish})
                in [Index = [Index] + 1, MaxFinish = NextMax], each [MaxFinish])),
            PreviousID = Table.AddColumn(Indexed, "PreviousSourceShiftID", each if [EmployeeShiftIndex] = 0 then null else ShiftIDs{[EmployeeShiftIndex] - 1}, type nullable text),
            PreviousFinish = Table.AddColumn(PreviousID, "PreviousShiftFinish", each if [EmployeeShiftIndex] = 0 then null else Finishes{[EmployeeShiftIndex] - 1}, type nullable datetime),
            PreviousDuty = Table.AddColumn(PreviousFinish, "PreviousDutyStatus", each if [EmployeeShiftIndex] = 0 then "Unknown" else DutyValues{[EmployeeShiftIndex] - 1}, type text),
            NextStart = Table.AddColumn(PreviousDuty, "NextShiftStart", each if [EmployeeShiftIndex] + 1 >= List.Count(Starts) then null else Starts{[EmployeeShiftIndex] + 1}, type nullable datetime),
            GapHours = Table.AddColumn(NextStart, "GapHours", each if [PreviousShiftFinish] = null or [ShiftStart] = null then null else Duration.TotalHours([ShiftStart] - [PreviousShiftFinish]), type nullable number),
            OverlapStatus = Table.AddColumn(GapHours, "OverlapStatus", each if [ShiftStart] = null or [ShiftFinish] = null then "Unknown"
                else if (PriorMaxFinishes{[EmployeeShiftIndex]} <> null and PriorMaxFinishes{[EmployeeShiftIndex]} > [ShiftStart]) or ([NextShiftStart] <> null and [NextShiftStart] < [ShiftFinish]) then "Overlap" else "Clear", type text)
        in OverlapStatus,
    EmployeeGroups = Table.Group(CoverageStatus, {"EmployeeKey"}, {{"EmployeeShifts", PrepareChronology, type table}}),
    ChronologicalShifts = if Table.IsEmpty(CoverageStatus) then PrepareChronology(CoverageStatus) else Table.Combine(EmployeeGroups[EmployeeShifts]),
    // Local datetime alone cannot establish elapsed hours across Sydney clock-change windows.
    GapWallHours = Table.RenameColumns(ChronologicalShifts, {{"GapHours", "GapWallHours"}}),
    GapElapsedStatus = Table.AddColumn(GapWallHours, "GapElapsedStatus", each if [PreviousShiftFinish] = null or [ShiftStart] = null then "Unknown"
        else if fnOvertimeRosterChecksClockSafe(List.Min({[PreviousShiftFinish], [ShiftStart]}), List.Max({[PreviousShiftFinish], [ShiftStart]})) then "Known" else "Unknown", type text),
    GapHours = Table.AddColumn(GapElapsedStatus, "GapHours", each if [GapElapsedStatus] = "Known" then [GapWallHours] else null, type nullable number),
    GapStatus = Table.AddColumn(GapHours, "GapStatus", each if [GapHours] <> null and [DutyStatus] = "Known" and [PreviousDutyStatus] = "Known" and [IdentityStatus] = "Verified" and [MatchStatus] = "Unique" and [CoverageStatus] = "Declared"
        and [PreviousShiftFinish] >= [ExportCoverageStart] and [OverlapStatus] = "Clear" then "Known" else "Unknown", type text),
    MinimumBreakFlag = Table.AddColumn(GapStatus, "MinimumBreakFlag", each if [GapStatus] = "Known" and [RosterScopeStatus] = "Known" then [GapHours] < OvertimeRosterChecks_MinimumBreakHours else null, type nullable logical),
    MinBreakStatus = Table.AddColumn(MinimumBreakFlag, "MinBreakStatus", each if [MinimumBreakFlag] = null then "Unknown" else if [MinimumBreakFlag] then "Review" else "Clear", type text),
    ShortGapReviewFlag = Table.AddColumn(MinBreakStatus, "ShortGapReviewFlag", each if [EligibilityStatus] <> "Eligible" or [GapStatus] <> "Known" then null
        else if NormalizeKey([EmploymentType]) = "CASUAL" or [BrokenShiftException] = true then false
        else if [BrokenShiftException] = null then null else [GapHours] < OvertimeRosterChecks_ShortGapReviewHours, type nullable logical),
    ShortGapStatus = Table.AddColumn(ShortGapReviewFlag, "ShortGapStatus", each if [ShortGapReviewFlag] = null then "Unknown" else if [ShortGapReviewFlag] then "Review" else "Clear", type text),
    DailyExcessHours = Table.AddColumn(ShortGapStatus, "DailyExcessHours", each if [QualifyingNetHours] = null or [DailyHoursLimit] = null then null else List.Max({0, [QualifyingNetHours] - [DailyHoursLimit]}), type nullable number),
    DailyStatus = Table.AddColumn(DailyExcessHours, "DailyStatus", each if [EligibilityStatus] = "Exempt" or [EligibilityStatus] = "Outside scope" then "Not applicable"
        else if [EligibilityStatus] <> "Eligible" or [NetHoursStatus] <> "Known" or [MealTreatmentStatus] <> "Known" or [MatchStatus] <> "Unique" or [OverlapStatus] <> "Clear" or [CoverageStatus] <> "Declared" then "Unknown"
        else if [DailyExcessHours] = null then "Unknown" else if [DailyExcessHours] > 0 then "Exceeded" else "Clear", type text),
    DailyHoursFlag = Table.AddColumn(DailyStatus, "DailyHoursFlag", each if [DailyStatus] = "Exceeded" then true else if [DailyStatus] = "Clear" then false else null, type nullable logical),
    ClockStatus = Table.AddColumn(DailyHoursFlag, "ClockStatus", each if fnOvertimeRosterChecksClockSafe([ShiftStart], [ShiftFinish]) then "Known" else "Unknown", type text)
in
    ClockStatus;

shared OvertimeRosterChecks_WorkSegments = let
    Source = if OvertimeRosterChecks_MinimumFortnightDutyFreeDays <= OvertimeRosterChecks_FortnightLengthDays then OvertimeRosterChecks_ShiftPrepare
        else error "MinimumFortnightDutyFreeDays cannot exceed FortnightLengthDays.",
    SegmentLists = Table.AddColumn(Source, "WorkSegments", each
        let ShiftID = [SourceShiftID], Employee = [EmployeeKey], WorkHours = [QualifyingNetHours],
            Start = [ShiftStart], Finish = [ShiftFinish], MealStart = [MealStartDateTime], MealFinish = [MealFinishDateTime],
            ClockEqualsNet = WorkHours <> null and [ClockStatus] = "Known" and [GrossHours] <> null and Number.Abs(WorkHours - [GrossHours]) < 0.0000001,
            UnpaidMealKnown = WorkHours <> null and [ClockStatus] = "Known" and [MealTreatmentStatus] = "Known" and [MealAvailabilityRequired] = false and MealStart <> null and MealFinish <> null
                and Start <> null and Finish <> null and MealStart >= Start and MealFinish <= Finish and MealFinish > MealStart,
            MealReconciles = if not UnpaidMealKnown then false else Number.Abs(Duration.TotalHours(Finish - Start) - Duration.TotalHours(MealFinish - MealStart) - WorkHours) < 0.0000001,
            FirstHours = if MealReconciles then Duration.TotalHours(MealStart - Start) else null,
            Rows = if WorkHours = null then {[SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = null, WorkFinish = null, NetWorkStartHours = 0, NetWorkEndHours = null, TimingStatus = "Unknown"]}
                else if ClockEqualsNet then {[SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = Start, WorkFinish = Finish, NetWorkStartHours = 0, NetWorkEndHours = WorkHours, TimingStatus = "Known"]}
                else if MealReconciles then {
                    [SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = Start, WorkFinish = MealStart, NetWorkStartHours = 0, NetWorkEndHours = FirstHours, TimingStatus = "Known"],
                    [SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = MealFinish, WorkFinish = Finish, NetWorkStartHours = FirstHours, NetWorkEndHours = WorkHours, TimingStatus = "Known"]}
                else {[SourceShiftID = ShiftID, EmployeeKey = Employee, WorkStart = null, WorkFinish = null, NetWorkStartHours = 0, NetWorkEndHours = WorkHours, TimingStatus = "Unknown"]}
        in Rows, type list),
    CombinedSegments = if Table.IsEmpty(Source) then {} else List.Combine(SegmentLists[WorkSegments]),
    SegmentTable = Table.FromRecords(CombinedSegments, type table [SourceShiftID = text, EmployeeKey = text, WorkStart = nullable datetime, WorkFinish = nullable datetime, NetWorkStartHours = number, NetWorkEndHours = nullable number, TimingStatus = text]),
    NonzeroSegments = Table.SelectRows(SegmentTable, each [NetWorkEndHours] = null or [NetWorkEndHours] > [NetWorkStartHours])
in
    NonzeroSegments;

shared OvertimeRosterChecks_FortnightParts = let
    Source = OvertimeRosterChecks_ShiftPrepare,
    Anchor = OvertimeRosterChecks_FortnightStartDate,
    FortnightDays = if OvertimeRosterChecks_MinimumFortnightDutyFreeDays <= OvertimeRosterChecks_FortnightLengthDays then OvertimeRosterChecks_FortnightLengthDays
        else error "MinimumFortnightDutyFreeDays cannot exceed FortnightLengthDays.",
    WorkSegments = Table.Buffer(OvertimeRosterChecks_WorkSegments),
    ShiftWorkJoin = Table.NestedJoin(Source, {"SourceShiftID"}, WorkSegments, {"SourceShiftID"}, "ShiftWork", JoinKind.LeftOuter),
    PartLists = Table.AddColumn(ShiftWorkJoin, "FortnightParts", each
        if Anchor = null or [ShiftStart] = null or [ShiftFinish] = null or [ShiftFinish] <= [ShiftStart] then {} else
        let Duty = _, ShiftID = [SourceShiftID], Start = [ShiftStart], Finish = [ShiftFinish],
            FirstIndex = Number.RoundDown(Duration.Days(Date.From(Start) - Anchor) / FortnightDays),
            LastIndex = Number.RoundDown(Duration.Days(Date.From(Finish) - Anchor) / FortnightDays),
            Indices = {FirstIndex..LastIndex},
            ShiftSegments = [ShiftWork],
            TimingKnown = not Table.IsEmpty(ShiftSegments) and List.AllTrue(List.Transform(ShiftSegments[TimingStatus], each _ = "Known")),
            Rows = List.Transform(Indices, (index) =>
                let FNStart = Date.AddDays(Anchor, index * FortnightDays), FNEnd = Date.AddDays(FNStart, FortnightDays),
                    PartStart = List.Max({Start, DateTime.From(FNStart)}), PartFinish = List.Min({Finish, DateTime.From(FNEnd)}),
                    IsWholeShift = PartStart = Start and PartFinish = Finish,
                    HoursBeforePart = if IsWholeShift then 0 else if not TimingKnown then null else List.Sum(List.Transform(Table.ToRecords(ShiftSegments), (segment) =>
                        if segment[WorkStart] >= PartStart then 0 else Duration.TotalHours(List.Min({segment[WorkFinish], PartStart}) - segment[WorkStart]))),
                    NetPartHours = if PartFinish <= PartStart then 0 else if IsWholeShift then Duty[QualifyingNetHours]
                        else if not TimingKnown then null else List.Sum(List.Transform(Table.ToRecords(ShiftSegments), (segment) =>
                            let Left = List.Max({segment[WorkStart], PartStart}), Right = List.Min({segment[WorkFinish], PartFinish})
                            in if Right > Left then Duration.TotalHours(Right - Left) else 0)),
                    PartNetEnd = if NetPartHours = null or HoursBeforePart = null then null else HoursBeforePart + NetPartHours
                in [SourceShiftID = ShiftID, EmployeeKey = Duty[EmployeeKey], EmployeeID = Duty[EmployeeID], EmployerKey = Duty[EmployerKey],
                    PayrollFortnightStart = FNStart, PayrollFortnightEnd = FNEnd, PartStart = PartStart, PartFinish = PartFinish,
                    PartNetStartHours = HoursBeforePart, PartNetEndHours = PartNetEnd, NetPartHours = NetPartHours,
                    DailyHoursLimit = Duty[DailyHoursLimit], EligibilityStatus = Duty[EligibilityStatus], RosterScopeStatus = Duty[RosterScopeStatus], DutyStatus = Duty[DutyStatus], IdentityStatus = Duty[IdentityStatus], MatchStatus = Duty[MatchStatus],
                    MealTreatmentStatus = Duty[MealTreatmentStatus], CoverageStatus = Duty[CoverageStatus], ExportCoverageStart = Duty[ExportCoverageStart], ExportCoverageEnd = Duty[ExportCoverageEnd], OverlapStatus = Duty[OverlapStatus]])
        in Rows, type list),
    CombinedParts = if Table.IsEmpty(Source) then {} else List.Combine(PartLists[FortnightParts]),
    PartTable = Table.FromRecords(CombinedParts, type table [SourceShiftID = text, EmployeeKey = text, EmployeeID = nullable text, EmployerKey = nullable text,
        PayrollFortnightStart = date, PayrollFortnightEnd = date, PartStart = datetime, PartFinish = datetime, PartNetStartHours = nullable number, PartNetEndHours = nullable number, NetPartHours = nullable number,
        DailyHoursLimit = nullable number, EligibilityStatus = text, RosterScopeStatus = text, DutyStatus = text, IdentityStatus = text, MatchStatus = text, MealTreatmentStatus = text, CoverageStatus = text,
        ExportCoverageStart = nullable datetime, ExportCoverageEnd = nullable datetime, OverlapStatus = text]),
    PositiveParts = Table.SelectRows(PartTable, each [PartFinish] > [PartStart]),
    AddCounter = (parts as table) as table =>
        let Ordered = Table.Sort(parts, {{"PartStart", Order.Ascending}, {"SourceShiftID", Order.Ascending}}),
            Indexed = Table.AddIndexColumn(Ordered, "PartIndex", 0, 1, Int64.Type),
            Hours = List.Buffer(Indexed[NetPartHours]), Count = List.Count(Hours),
            BeforeValues = List.Buffer(List.Generate(() => [Index = 0, Before = 0], each [Index] < Count,
                each let NextIndex = [Index] + 1, CurrentHours = Hours{[Index]}, NextBefore = if [Before] = null or CurrentHours = null then null else [Before] + CurrentHours
                in [Index = NextIndex, Before = NextBefore], each [Before])),
            Before = Table.AddColumn(Indexed, "FortnightHoursBefore", each BeforeValues{[PartIndex]}, type nullable number),
            After = Table.AddColumn(Before, "FortnightHoursAfter", each if [FortnightHoursBefore] = null or [NetPartHours] = null then null else [FortnightHoursBefore] + [NetPartHours], type nullable number),
            Excess = Table.AddColumn(After, "FortnightExcessHours", each if [FortnightHoursBefore] = null or [FortnightHoursAfter] = null then null
                else List.Max({0, [FortnightHoursAfter] - OvertimeRosterChecks_FortnightHoursLimit}) - List.Max({0, [FortnightHoursBefore] - OvertimeRosterChecks_FortnightHoursLimit}), type nullable number),
            RemovedIndex = Table.RemoveColumns(Excess, {"PartIndex"})
        in RemovedIndex,
    EmployeeFortnights = Table.Group(PositiveParts, {"EmployeeKey", "PayrollFortnightStart"}, {{"Parts", AddCounter, type table}}),
    CountedParts = if Table.IsEmpty(PositiveParts) then AddCounter(PositiveParts) else Table.Combine(EmployeeFortnights[Parts])
in
    CountedParts;

shared OvertimeRosterChecks_FortnightResults = let
    Parts = OvertimeRosterChecks_FortnightParts,
    DutyColumns = Table.SelectColumns(OvertimeRosterChecks_ShiftPrepare, {"EmployeeKey", "SourceDate", "ShiftStart", "ShiftFinish", "IdentityStatus", "MatchStatus", "DutyStatus", "RosterScopeStatus"}),
    Duties = Table.Buffer(DutyColumns),
    GroupedParts = Table.Group(Parts, {"EmployeeKey", "PayrollFortnightStart", "PayrollFortnightEnd"}, {{"Parts", each _, type table}}),
    EmployerKey = Table.AddColumn(GroupedParts, "EmployerKey", each [Parts]{0}[EmployerKey], type nullable text),
    EmployeeID = Table.AddColumn(EmployerKey, "EmployeeID", each [Parts]{0}[EmployeeID], type nullable text),
    NetRosterHours = Table.AddColumn(EmployeeID, "NetRosterHours", each if List.NonNullCount([Parts][NetPartHours]) = Table.RowCount([Parts]) then List.Sum([Parts][NetPartHours]) else null, type nullable number),
    ExcessHours = Table.AddColumn(NetRosterHours, "FortnightExcessHours", each if [NetRosterHours] = null then null else List.Max({0, [NetRosterHours] - OvertimeRosterChecks_FortnightHoursLimit}), type nullable number),
    CoverageStatus = Table.AddColumn(ExcessHours, "CoverageStatus", each
        let Row = [Parts]{0}, Start = DateTime.From([PayrollFortnightStart]), Finish = DateTime.From([PayrollFortnightEnd])
        in if Row[ExportCoverageStart] <> null and Row[ExportCoverageEnd] <> null and Row[ExportCoverageStart] <= Start and Row[ExportCoverageEnd] >= Finish then "Declared" else "Unknown", type text),
    EmployeeDutyJoin = Table.NestedJoin(CoverageStatus, {"EmployeeKey"}, Duties, {"EmployeeKey"}, "EmployeeDutyRows", JoinKind.LeftOuter),
    EmployeeDuty = Table.AddColumn(EmployeeDutyJoin, "Duty", each
        let StartDate = [PayrollFortnightStart], FinishDate = [PayrollFortnightEnd], Start = DateTime.From(StartDate), Finish = DateTime.From(FinishDate)
        in Table.SelectRows([EmployeeDutyRows], each
            let ValidSpan = [ShiftStart] <> null and [ShiftFinish] <> null and [ShiftFinish] > [ShiftStart]
            in if ValidSpan then [ShiftStart] < Finish and [ShiftFinish] > Start
                else if [SourceDate] <> null then [SourceDate] >= StartDate and [SourceDate] < FinishDate
                else true), type table),
    DutyKnown = Table.AddColumn(EmployeeDuty, "DutyCoverageKnown", each [CoverageStatus] = "Declared" and not Table.IsEmpty([Duty])
        and List.AllTrue(List.Transform(Table.ToRecords([Duty]), each [IdentityStatus] = "Verified" and [MatchStatus] = "Unique" and [DutyStatus] = "Known")), type logical),
    DutyFreeDays = Table.AddColumn(DutyKnown, "DutyFreeDays", each if not [DutyCoverageKnown] then null else
        let Days = List.Dates([PayrollFortnightStart], Duration.Days([PayrollFortnightEnd] - [PayrollFortnightStart]), #duration(1, 0, 0, 0)), Duty = [Duty],
            FreeDays = List.Select(Days, (day) =>
                let Start = DateTime.From(day), Finish = DateTime.From(Date.AddDays(day, 1)), Overlaps = Table.SelectRows(Duty, each [ShiftStart] < Finish and [ShiftFinish] > Start)
                in Table.IsEmpty(Overlaps))
        in List.Count(FreeDays), type nullable number),
    DaysShort = Table.AddColumn(DutyFreeDays, "DaysShort", each if [DutyFreeDays] = null then null else List.Max({0, OvertimeRosterChecks_MinimumFortnightDutyFreeDays - [DutyFreeDays]}), type nullable number),
    DaysOffFlag = Table.AddColumn(DaysShort, "DaysOffFlag", each if [DaysShort] = null or not List.AnyTrue(List.Transform([Duty][RosterScopeStatus], each _ = "Known")) then null else [DaysShort] > 0, type nullable logical),
    DaysOffStatus = Table.AddColumn(DaysOffFlag, "DaysOffStatus", each if [DaysOffFlag] = null then "Unknown" else if [DaysOffFlag] then "Review" else "Clear", type text),
    FortnightStatus = Table.AddColumn(DaysOffStatus, "FortnightStatus", each if not [DutyCoverageKnown] or [NetRosterHours] = null
        or not List.AllTrue(List.Transform(Table.ToRecords([Parts]), each [IdentityStatus] = "Verified" and [MatchStatus] = "Unique" and [MealTreatmentStatus] = "Known" and [DutyStatus] = "Known" and [OverlapStatus] = "Clear")) then "Unknown"
        else if [FortnightExcessHours] > 0 then "Exceeded" else "Clear", type text),
    PublishedColumns = Table.SelectColumns(FortnightStatus, {"EmployeeKey", "EmployeeID", "EmployerKey", "PayrollFortnightStart", "PayrollFortnightEnd", "NetRosterHours", "FortnightExcessHours", "DutyFreeDays", "DaysShort", "DaysOffFlag", "CoverageStatus", "FortnightStatus", "DaysOffStatus"})
in
    PublishedColumns;

shared OvertimeRosterChecks_ExcessSegments = let
    // Keep provenance and unrelated check columns out of the per-part excess calculation.
    ShiftColumns = Table.SelectColumns(OvertimeRosterChecks_ShiftPrepare, {"SourceShiftID", "EmployeeKey", "QualifyingNetHours", "DailyHoursLimit", "EligibilityStatus", "MealTreatmentStatus", "MatchStatus", "DutyStatus", "CoverageStatus", "OverlapStatus", "DailyStatus"}),
    Parts = OvertimeRosterChecks_FortnightParts,
    // Materialise only scalar keys/status before nested joins revisit each shift's parts.
    FNStatusColumns = Table.SelectColumns(OvertimeRosterChecks_FortnightResults, {"EmployeeKey", "PayrollFortnightStart", "FortnightStatus"}),
    BufferedFNStatus = Table.Buffer(FNStatusColumns),
    FNStatusJoin = Table.NestedJoin(Parts, {"EmployeeKey", "PayrollFortnightStart"}, BufferedFNStatus, {"EmployeeKey", "PayrollFortnightStart"}, "FNStatus", JoinKind.LeftOuter),
    FNStatus = Table.ExpandTableColumn(FNStatusJoin, "FNStatus", {"FortnightStatus"}, {"PartFortnightStatus"}),
    PartsColumns = Table.SelectColumns(FNStatus, {"SourceShiftID", "EmployeeKey", "PayrollFortnightStart", "PartNetStartHours", "PartNetEndHours", "FortnightHoursBefore", "PartFortnightStatus"}),
    BufferedParts = Table.Buffer(PartsColumns),
    ShiftPartJoin = Table.NestedJoin(ShiftColumns, {"SourceShiftID"}, BufferedParts, {"SourceShiftID"}, "Parts", JoinKind.LeftOuter),
    // If fortnight splitting cannot locate net work, retain whole-shift daily positions without inventing a period allocation.
    PositionParts = Table.AddColumn(ShiftPartJoin, "PositionParts", each if Table.IsEmpty([Parts])
        or List.NonNullCount([Parts][PartNetStartHours]) <> Table.RowCount([Parts]) or List.NonNullCount([Parts][PartNetEndHours]) <> Table.RowCount([Parts]) then
        #table({"PayrollFortnightStart", "PartNetStartHours", "PartNetEndHours", "FortnightHoursBefore", "PartFortnightStatus"}, {{null, 0, [QualifyingNetHours], null, "Unknown"}})
        else Table.SelectColumns([Parts], {"PayrollFortnightStart", "PartNetStartHours", "PartNetEndHours", "FortnightHoursBefore", "PartFortnightStatus"}), type table),
    ExpandedPositionParts = Table.ExpandTableColumn(PositionParts, "PositionParts", {"PayrollFortnightStart", "PartNetStartHours", "PartNetEndHours", "FortnightHoursBefore", "PartFortnightStatus"}),
    ExpandedParts = Table.RemoveColumns(ExpandedPositionParts, {"Parts"}),
    WorkSegments = Table.Buffer(OvertimeRosterChecks_WorkSegments),
    WorkJoin = Table.NestedJoin(ExpandedParts, {"SourceShiftID"}, WorkSegments, {"SourceShiftID"}, "ShiftWork", JoinKind.LeftOuter),
    SegmentLists = Table.AddColumn(WorkJoin, "ExcessSegments", each
        let Duty = _, ShiftID = [SourceShiftID], Left = [PartNetStartHours], Right = [PartNetEndHours], DailyLimit = [DailyHoursLimit],
            FNBefore = [FortnightHoursBefore], FNLimitPosition = if FNBefore = null or Left = null then null else Left + List.Max({0, OvertimeRosterChecks_FortnightHoursLimit - FNBefore}),
            ThisWork = [ShiftWork],
            CandidateCuts = List.Combine({{Left, Right, DailyLimit, FNLimitPosition}, ThisWork[NetWorkStartHours], ThisWork[NetWorkEndHours]}),
            Cuts = if Left = null or Right = null then {} else List.Sort(List.Distinct(List.Select(List.RemoveNulls(CandidateCuts), each _ >= Left and _ <= Right))),
            Pairs = if List.Count(Cuts) < 2 then {} else List.Transform({0..List.Count(Cuts) - 2}, each [Left = Cuts{_}, Right = Cuts{_ + 1}]),
            LocateClock = (position as number, isFinish as logical) as nullable datetime =>
                let Matching = Table.SelectRows(ThisWork, each [TimingStatus] = "Known" and (if isFinish then position > [NetWorkStartHours] and position <= [NetWorkEndHours] else position >= [NetWorkStartHours] and position < [NetWorkEndHours])),
                    Located = if Table.RowCount(Matching) = 1 then Matching{0}[WorkStart] + #duration(0, 0, 0, (position - Matching{0}[NetWorkStartHours]) * 3600) else null
                in Located,
            Rows = List.Transform(Pairs, (pair) =>
                let Daily = DailyLimit <> null and pair[Left] >= DailyLimit, FN = FNLimitPosition <> null and pair[Left] >= FNLimitPosition,
                    Start = LocateClock(pair[Left], false), Finish = LocateClock(pair[Right], true),
                    Triggers = Text.Combine(List.RemoveNulls({if Daily then "Daily" else null, if FN then "Fortnight76" else null}), "; "),
                    Status = if Duty[EligibilityStatus] = "Eligible" and Duty[MealTreatmentStatus] = "Known" and Duty[MatchStatus] = "Unique" and Duty[DutyStatus] = "Known" and Duty[CoverageStatus] = "Declared" and Duty[OverlapStatus] = "Clear"
                        and (not Daily or Duty[DailyStatus] <> "Unknown") and (not FN or Duty[PartFortnightStatus] <> "Unknown") then "Review" else "Unknown"
                in [SourceShiftID = ShiftID, EmployeeKey = Duty[EmployeeKey], PayrollFortnightStart = Duty[PayrollFortnightStart], NetWorkStartHours = pair[Left], NetWorkEndHours = pair[Right], SegmentHours = pair[Right] - pair[Left],
                    SegmentStart = Start, SegmentFinish = Finish, DailyTrigger = Daily, FortnightTrigger = FN, TriggerSummary = Triggers, TimingStatus = if Start <> null and Finish <> null then "Known" else "Unknown", OvertimeStatus = Status])
        in List.Select(Rows, each [SegmentHours] > 0 and ([DailyTrigger] or [FortnightTrigger])), type list),
    CombinedSegments = if Table.IsEmpty(ExpandedParts) then {} else List.Combine(SegmentLists[ExcessSegments]),
    SegmentTable = Table.FromRecords(CombinedSegments, type table [SourceShiftID = text, EmployeeKey = text, PayrollFortnightStart = nullable date, NetWorkStartHours = number, NetWorkEndHours = number, SegmentHours = number,
        SegmentStart = nullable datetime, SegmentFinish = nullable datetime, DailyTrigger = logical, FortnightTrigger = logical, TriggerSummary = text, TimingStatus = text, OvertimeStatus = text])
in
    SegmentTable;

shared OvertimeRosterChecks_ShiftResults = let
    Source = OvertimeRosterChecks_ShiftPrepare,
    Parts = OvertimeRosterChecks_FortnightParts,
    FNJoin = Table.NestedJoin(Parts, {"EmployeeKey", "PayrollFortnightStart"}, OvertimeRosterChecks_FortnightResults, {"EmployeeKey", "PayrollFortnightStart"}, "FN", JoinKind.LeftOuter),
    ExpandedFN = Table.ExpandTableColumn(FNJoin, "FN", {"DutyFreeDays", "DaysShort", "DaysOffFlag", "DaysOffStatus", "FortnightStatus"}),
    ShiftPartJoin = Table.NestedJoin(Source, {"SourceShiftID"}, ExpandedFN, {"SourceShiftID"}, "ShiftParts", JoinKind.LeftOuter),
    PayrollFortnightCount = Table.AddColumn(ShiftPartJoin, "PayrollFortnightCount", each Table.RowCount([ShiftParts]), Int64.Type),
    PayrollFortnightStart = Table.AddColumn(PayrollFortnightCount, "PayrollFortnightStart", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[PayrollFortnightStart] else null, type nullable date),
    PayrollFortnightEnd = Table.AddColumn(PayrollFortnightStart, "PayrollFortnightEnd", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[PayrollFortnightEnd] else null, type nullable date),
    FortnightHoursBefore = Table.AddColumn(PayrollFortnightEnd, "FortnightHoursBefore", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[FortnightHoursBefore] else null, type nullable number),
    FortnightHoursAfter = Table.AddColumn(FortnightHoursBefore, "FortnightHoursAfter", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[FortnightHoursAfter] else null, type nullable number),
    FortnightExcessHours = Table.AddColumn(FortnightHoursAfter, "FortnightExcessHours", each if [PayrollFortnightCount] = 0 or List.NonNullCount([ShiftParts][FortnightExcessHours]) <> [PayrollFortnightCount] then null else List.Sum([ShiftParts][FortnightExcessHours]), type nullable number),
    FortnightStatus = Table.AddColumn(FortnightExcessHours, "FortnightStatus", each if [PayrollFortnightCount] = 0 or List.Contains([ShiftParts][FortnightStatus], "Unknown") then "Unknown"
        else if [FortnightExcessHours] > 0 then "Exceeded" else "Clear", type text),
    FortnightHoursFlag = Table.AddColumn(FortnightStatus, "FortnightHoursFlag", each if [EligibilityStatus] <> "Eligible" or [FortnightStatus] = "Unknown" then null else [FortnightExcessHours] > 0, type nullable logical),
    DutyFreeDays = Table.AddColumn(FortnightHoursFlag, "DutyFreeDays", each if [PayrollFortnightCount] = 1 then [ShiftParts]{0}[DutyFreeDays] else null, type nullable number),
    DaysShort = Table.AddColumn(DutyFreeDays, "DaysShort", each if [PayrollFortnightCount] = 0 or List.NonNullCount([ShiftParts][DaysShort]) <> [PayrollFortnightCount] then null else List.Max([ShiftParts][DaysShort]), type nullable number),
    DaysOffFlag = Table.AddColumn(DaysShort, "DaysOffFlag", each if [PayrollFortnightCount] = 0 then null else if List.Contains([ShiftParts][DaysOffFlag], true) then true
        else if List.Contains([ShiftParts][DaysOffFlag], null) then null else false, type nullable logical),
    DaysOffStatus = Table.AddColumn(DaysOffFlag, "DaysOffStatus", each if [DaysOffFlag] = null then "Unknown" else if [DaysOffFlag] then "Review" else "Clear", type text),
    DaysOffOvertimeReviewFlag = Table.AddColumn(DaysOffStatus, "DaysOffOvertimeReviewFlag", each if [EligibilityStatus] <> "Eligible" or [DaysOffFlag] = null then null
        else if Text.Upper(Text.Remove([EmploymentType], {" ", "-", "_"})) = "CASUAL" then false else [DaysOffFlag], type nullable logical),
    WorkTiming = Table.Group(OvertimeRosterChecks_WorkSegments, {"SourceShiftID"}, {{"MealTimingStatus", each if List.AllTrue(List.Transform([TimingStatus], each _ = "Known")) then "Known" else "Unknown", type text}}),
    WorkTimingJoin = Table.NestedJoin(DaysOffOvertimeReviewFlag, {"SourceShiftID"}, WorkTiming, {"SourceShiftID"}, "WorkTiming", JoinKind.LeftOuter),
    MealTimingStatus = Table.ExpandTableColumn(WorkTimingJoin, "WorkTiming", {"MealTimingStatus"}),
    TimingStatusDefaults = Table.TransformColumns(MealTimingStatus, {{"MealTimingStatus", each if _ = null then "Unknown" else _, type text}}),
    OvertimeJoin = Table.NestedJoin(TimingStatusDefaults, {"SourceShiftID"}, OvertimeRosterChecks_ExcessSegments, {"SourceShiftID"}, "ExcessSegments", JoinKind.LeftOuter),
    OvertimeHours = Table.AddColumn(OvertimeJoin, "OvertimeHours", each if [QualifyingNetHours] = null or [DailyHoursLimit] = null or [FortnightStatus] = "Unknown" then null else List.Sum({0} & [ExcessSegments][SegmentHours]), type nullable number),
    OvertimeStatus = Table.AddColumn(OvertimeHours, "OvertimeStatus", each if [EligibilityStatus] = "Exempt" or [EligibilityStatus] = "Outside scope" then "Not applicable"
        else if [EligibilityStatus] <> "Eligible" or [DailyStatus] = "Unknown" or [FortnightStatus] = "Unknown" then "Unknown"
        else if [OvertimeHours] > 0 or [ShortGapReviewFlag] = true or [DaysOffOvertimeReviewFlag] = true then "Review" else "Clear", type text),
    AnyRuleFlag = Table.AddColumn(OvertimeStatus, "AnyRuleFlag", each
        let Flags = {[DailyHoursFlag], [MinimumBreakFlag], [ShortGapReviewFlag], [DaysOffFlag], [FortnightHoursFlag]}
        in if List.Contains(Flags, true) then true else if List.Contains(Flags, null) then null else false, type nullable logical),
    TriggerSummary = Table.AddColumn(AnyRuleFlag, "TriggerSummary", each Text.Combine(List.RemoveNulls({if [DailyHoursFlag] = true then "Daily" else null,
        if [MinimumBreakFlag] = true then "MinimumBreak" else null, if [ShortGapReviewFlag] = true then "ShortGapReview" else null,
        if [DaysOffFlag] = true then "DaysOffRoster" else null, if [DaysOffOvertimeReviewFlag] = true then "DaysOffOvertimeReview" else null,
        if [FortnightHoursFlag] = true then "Fortnight76" else null}), "; "), type text),
    ResultReasons = Table.AddColumn(TriggerSummary, "ResultReasons", each Text.Combine(List.RemoveNulls({
        if [IdentityStatus] <> "Verified" then "Employee, allocation-alias or facility identity unresolved" else null,
        if [MatchStatus] <> "Unique" then "Shift match or conflicting duplicate unresolved" else null,
        if [SourceRecordCount] > 1 and [SourceFactCount] = 1 then "Identical published export copies collapsed" else null,
        if [PublicationStatus] <> "Known" then "Publication flag missing" else null,
        if [DutyStatus] <> "Known" then "Duty status unresolved" else null,
        if [PolicyStatus] <> "Known" then "Roster and worker policy facts conflict" else null,
        if [CoverageStatus] <> "Declared" then "Export coverage is not declared across all configured units" else null,
        if [AgreementStream] = null then "Agreement stream missing" else null,
        if [Grade] = null then "Grade missing" else null,
        if [OvertimeExempt] = null then "Overtime exemption unresolved" else null,
        if [BrokenShiftException] = null then "Broken-shift exception unresolved" else null,
        if [NetHoursStatus] <> "Known" then "Net roster hours or shift span invalid or conflicting" else null,
        if [MealTreatmentStatus] <> "Known" then "Meal availability treatment unresolved" else null,
        if [MealTimingStatus] <> "Known" then "Exact meal or overtime clock timing unresolved" else null,
        if [PreviousSourceShiftID] = null then "First supplied roster shift has no preceding shift" else null,
        if [GapElapsedStatus] <> "Known" and [PreviousSourceShiftID] <> null then "Elapsed gap cannot be established across local clock ambiguity" else null,
        if [OverlapStatus] = "Overlap" then "Overlapping roster shifts require reconciliation" else null,
        if OvertimeRosterChecks_FortnightStartDate = null then "Fortnight start is unknown because Settings permutation dates are empty" else null,
        if [FortnightStatus] = "Unknown" and OvertimeRosterChecks_FortnightStartDate <> null then "Roster-based fortnight has incomplete coverage or unresolved qualifying hours" else null}), "; "), type text),
    PublishedColumns = Table.SelectColumns(ResultReasons, {"SourceShiftID", "SourceUnit", "SourceExportUnit", "SourceUnits", "SourceFile", "SourceRow", "SourceName", "SourceRole", "SourceDate", "EmployerKey", "WorkerEmployeeID", "AllocationCode", "EmployeeID", "EmployeeKey", "Facility", "Role", "ShiftStart", "ShiftFinish", "RosterShiftPeriod", "GrossHours", "BreakMinutes", "NetRosterHours", "QualifyingNetHours", "EmploymentType", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException", "IdentityStatus", "MatchStatus", "PolicyStatus", "PublicationStatus", "DutyStatus", "RosterScopeStatus", "EligibilityStatus", "CoverageStatus", "NetHoursStatus", "MealTreatmentStatus", "MealTimingStatus", "ClockStatus", "PreviousSourceShiftID", "PreviousShiftFinish", "GapWallHours", "GapHours", "GapElapsedStatus", "GapStatus", "OverlapStatus", "MinimumBreakFlag", "MinBreakStatus", "ShortGapReviewFlag", "ShortGapStatus", "DailyHoursLimit", "DailyExcessHours", "DailyHoursFlag", "DailyStatus", "PayrollFortnightCount", "PayrollFortnightStart", "PayrollFortnightEnd", "FortnightHoursBefore", "FortnightHoursAfter", "FortnightExcessHours", "FortnightHoursFlag", "FortnightStatus", "DutyFreeDays", "DaysShort", "DaysOffFlag", "DaysOffStatus", "DaysOffOvertimeReviewFlag", "OvertimeHours", "OvertimeStatus", "AnyRuleFlag", "TriggerSummary", "ResultReasons"}),
    OrderedOutput = Table.Sort(PublishedColumns, {{"SourceUnit", Order.Ascending}, {"EmployeeKey", Order.Ascending}, {"ShiftStart", Order.Ascending}, {"SourceShiftID", Order.Ascending}})
in
    OrderedOutput;