section OvertimeRosterChecksNative;
shared OvertimeRosterChecksNativeTests = () as table => let
Evaluate = (Inputs as record) as record => let
#"IMPORT OvertimeRosterChecks CentriSyncPaths" = Inputs[PathMapping],
#"INPUT OvertimeRosterChecks FilePathUrl" = Inputs[PathInput],
OvertimeRosterChecks_Path = let
    NormalizePath = (value as any) as nullable text =>
        if value = null then null else Text.TrimEnd(Text.Replace(Text.Trim(Text.From(value)), "/", "\\"), "\\"),
    PathInput = #"INPUT OvertimeRosterChecks FilePathUrl",
    PathColumn = if Table.HasColumns(PathInput, {"FilePath"}) then PathInput[FilePath]
        else if Table.ColumnNames(PathInput) = {"Column1"} then PathInput[Column1]
        else error "FilePathUrl must have a FilePath column or be one named cell.",
    RawPath = if List.Count(PathColumn) = 1 then NormalizePath(PathColumn{0}) else error "FilePathUrl requires exactly one value.",
    NonBlankPath = if RawPath = null or RawPath = "" then error "Save and recalculate OvertimeRosterChecks.xlsx before resolving FilePathUrl." else RawPath,
    WorkbookPath = if Text.Contains(NonBlankPath, "[") then Text.BeforeDelimiter(NonBlankPath, "[") & Text.BetweenDelimiters(NonBlankPath, "[", "]") else NonBlankPath,
    FileName = Text.AfterDelimiter(WorkbookPath, "\\", {0, RelativePosition.FromEnd}),
    NamedWorkbookPath = if Comparer.OrdinalIgnoreCase(FileName, "OvertimeRosterChecks.xlsx") = 0 then WorkbookPath
        else error "FilePathUrl must identify OvertimeRosterChecks.xlsx.",
    MappingColumns = Table.SelectColumns(#"IMPORT OvertimeRosterChecks CentriSyncPaths", {"SharepointRootUrl", "SyncedFolderRootPath"}),
    NormalizedMappings = Table.TransformColumns(MappingColumns, {{"SharepointRootUrl", NormalizePath, type nullable text}, {"SyncedFolderRootPath", NormalizePath, type nullable text}}),
    BufferedMappings = Table.Buffer(NormalizedMappings),
    SharePointRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SharepointRootUrl], type nullable text),
    DocumentRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each if [SharepointRootUrl] = null then null else [SharepointRootUrl] & "\\Shared Documents", type nullable text),
    LocalRoots = Table.AddColumn(BufferedMappings, "MatchRoot", each [SyncedFolderRootPath], type nullable text),
    CandidateRoots = Table.Combine({SharePointRoots, DocumentRoots, LocalRoots}),
    RootLengths = Table.AddColumn(CandidateRoots, "RootLength", each if [MatchRoot] = null then 0 else Text.Length([MatchRoot]), Int64.Type),
    MatchingRoots = Table.SelectRows(RootLengths, each [RootLength] > 0 and [SyncedFolderRootPath] <> null and [SyncedFolderRootPath] <> ""
        and Text.StartsWith(NamedWorkbookPath, [MatchRoot], Comparer.OrdinalIgnoreCase)
        and (Text.Length(NamedWorkbookPath) = [RootLength] or Text.Range(NamedWorkbookPath, [RootLength], 1) = "\\")),
    OrderedRoots = Table.Sort(MatchingRoots, {{"RootLength", Order.Descending}}),
    LongestRoot = if Table.IsEmpty(OrderedRoots) then error "FilePathUrl has no CentriSyncPaths mapping." else OrderedRoots{0},
    EqualRoots = Table.SelectRows(OrderedRoots, each [RootLength] = LongestRoot[RootLength]),
    Destinations = List.Distinct(EqualRoots[SyncedFolderRootPath], Comparer.OrdinalIgnoreCase),
    SelectedRoot = if List.Count(Destinations) = 1 then LongestRoot else error "CentriSyncPaths has conflicting longest-prefix destinations.",
    RelativePath = Text.TrimStart(Text.Range(NamedWorkbookPath, SelectedRoot[RootLength]), "\\"),
    LocalWorkbookPath = SelectedRoot[SyncedFolderRootPath] & "\\" & RelativePath,
    WorkbookFolder = Text.BeforeDelimiter(LocalWorkbookPath, "\\", {0, RelativePosition.FromEnd}),
    NestedSuffix = "\\2. Calculations\\OvertimeRosterChecks",
    DirectSuffix = "\\2. Calculations",
    ClientPath = if Text.EndsWith(WorkbookFolder, NestedSuffix, Comparer.OrdinalIgnoreCase) then Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(NestedSuffix))
        else if Text.EndsWith(WorkbookFolder, DirectSuffix, Comparer.OrdinalIgnoreCase) then Text.Start(WorkbookFolder, Text.Length(WorkbookFolder) - Text.Length(DirectSuffix))
        else error "OvertimeRosterChecks.xlsx must be in the organisation's 2. Calculations folder or its OvertimeRosterChecks subfolder.",
    Result = [ClientPath = ClientPath, WorkbookPath = LocalWorkbookPath]
in
    Result,
fnOvertimeRosterChecksRead = (unit as text, relativeFile as text, item as text, kind as text) as table => error "Workbook connector helper is disabled in native fixtures.",
#"IMPORT BD Published Roster" = Inputs[RosterBD],
#"IMPORT TE Published Roster" = Inputs[RosterTE],
#"IMPORT JH-RY Published Roster" = Inputs[RosterJHRY],
#"IMPORT BD Workers" = Inputs[WorkersBD],
#"IMPORT TE Workers" = Inputs[WorkersTE],
#"IMPORT JH-RY Workers" = Inputs[WorkersJHRY],
#"IMPORT BD ShiftPeriod" = Inputs[PeriodsBD],
#"IMPORT TE ShiftPeriod" = Inputs[PeriodsTE],
#"IMPORT JH-RY ShiftPeriod" = Inputs[PeriodsJHRY],
#"IMPORT OvertimeRosterChecks Whiddon Lists" = error "SharePoint navigation is disabled in native fixtures.",
#"LINK OvertimeRosterChecks Roles" = Inputs[Roles],
#"LINK OvertimeRosterChecks Facilities" = Inputs[Facilities],
#"LINK OvertimeRosterThreshold" = Inputs[Thresholds],
fnOvertimeRosterChecksText = (value as any) as nullable text =>
let
    TextValue = if value = null then null else Text.Trim(Text.From(value, "en-AU")),
    Result = if TextValue = "" then null else TextValue
in
    Result,
fnOvertimeRosterChecksNumber = (value as any) as nullable number =>
let
    Parsed = try Number.From(value) otherwise null,
    Result = if Parsed = null then null else if Number.IsNaN(Parsed) or Number.Abs(Parsed) = #infinity then null else Parsed
in
    Result,
fnOvertimeRosterChecksLogical = (value as any) as nullable logical =>
let
    TextValue = fnOvertimeRosterChecksText(value),
    Key = if TextValue = null then null else Text.Upper(TextValue),
    Result = if Value.Is(value, type logical) then value else if List.Contains({"TRUE", "YES", "1"}, Key) then true
        else if List.Contains({"FALSE", "NO", "0"}, Key) then false else null
in
    Result,
fnOvertimeRosterChecksKey = (values as list) as text => Binary.ToText(Json.FromValue(values), BinaryEncoding.Base64),
fnOvertimeRosterChecksSheetPrepare = (source as table, unit as text, required as list) as table =>
let
    Headers = if Table.HasColumns(source, required) then source else Table.PromoteHeaders(source, [PromoteAllScalars = true]),
    MissingColumns = List.Difference(required, Table.ColumnNames(Headers)),
    RequiredSchema = if List.IsEmpty(MissingColumns) then Headers else error Error.Record("OvertimeRosterChecks source schema", "Required source columns are missing.", [SourceUnit = unit, MissingColumns = MissingColumns]),
    SourceRows = Table.AddIndexColumn(RequiredSchema, "SourceRow", 1, 1, Int64.Type),
    SourceUnit = Table.AddColumn(SourceRows, "SourceUnit", each unit, type text)
in
    SourceUnit,
OvertimeRosterChecks_DailyHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "DailyHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one DailyHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "DailyHoursLimit must be a positive number." else if ParsedValue <= 0 then error "DailyHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_CSHospitalityHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "CSHospitalityHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one CSHospitalityHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "CSHospitalityHoursLimit must be a positive number." else if ParsedValue <= 0 then error "CSHospitalityHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_CSPersonalCareHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "CSPersonalCareHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one CSPersonalCareHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "CSPersonalCareHoursLimit must be a positive number." else if ParsedValue <= 0 then error "CSPersonalCareHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_MinimumBreakHours = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "MinimumBreakHours"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one MinimumBreakHours row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "MinimumBreakHours must be a positive number." else if ParsedValue <= 0 then error "MinimumBreakHours has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_ShortGapReviewHours = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "ShortGapReviewHours"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one ShortGapReviewHours row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "ShortGapReviewHours must be a positive number." else if ParsedValue <= 0 then error "ShortGapReviewHours has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_MinimumFortnightDutyFreeDays = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "MinimumFortnightDutyFreeDays"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one MinimumFortnightDutyFreeDays row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "MinimumFortnightDutyFreeDays must be a positive number." else if ParsedValue <= 0 or Number.RoundDown(ParsedValue) <> ParsedValue then error "MinimumFortnightDutyFreeDays has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_FortnightHoursLimit = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "FortnightHoursLimit"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one FortnightHoursLimit row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "FortnightHoursLimit must be a positive number." else if ParsedValue <= 0 then error "FortnightHoursLimit has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_FortnightLengthDays = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "FortnightLengthDays"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one FortnightLengthDays row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then error "FortnightLengthDays must be a positive number." else if ParsedValue <= 0 or Number.RoundDown(ParsedValue) <> ParsedValue then error "FortnightLengthDays has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_MinimumDutyFreeDaysBeforeNight = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "MinimumDutyFreeDaysBeforeNight"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one MinimumDutyFreeDaysBeforeNight row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "Value", null),
    ParsedValue = fnOvertimeRosterChecksNumber(SuppliedValue),
    ValidatedValue = if ParsedValue = null then null else if ParsedValue < 0 then error "MinimumDutyFreeDaysBeforeNight has an invalid numeric value." else ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_PayrollFortnightAnchor = let
    Source = #"LINK OvertimeRosterThreshold",
    MatchingTitle = Table.SelectRows(Source, each [Title] = "PayrollFortnightAnchor"),
    SelectedRow = if Table.RowCount(MatchingTitle) = 1 then MatchingTitle{0} else error "OvertimeRosterThreshold requires exactly one PayrollFortnightAnchor row.",
    SuppliedValue = Record.FieldOrDefault(SelectedRow, "DateValue", null),
    ParsedValue = if SuppliedValue = null then null else try Date.From(SuppliedValue) otherwise error "PayrollFortnightAnchor DateValue must be a date.",
    ValidatedValue = ParsedValue
in
    ValidatedValue,
OvertimeRosterChecks_RosterPrepare = let
    RequiredColumns = {"Employee Code", "Employee_Code", "Employee Roster Name", "Role", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length"},
    BDRows = fnOvertimeRosterChecksSheetPrepare(#"IMPORT BD Published Roster", "BD", RequiredColumns),
    TERows = fnOvertimeRosterChecksSheetPrepare(#"IMPORT TE Published Roster", "TE", RequiredColumns),
    JHRYRows = fnOvertimeRosterChecksSheetPrepare(#"IMPORT JH-RY Published Roster", "JH-RY", RequiredColumns),
    CombinedRows = Table.Combine({BDRows, TERows, JHRYRows}),
    NeededColumns = Table.SelectColumns(CombinedRows, {"SourceUnit", "SourceRow", "Employee Code", "Employee_Code", "Pay Company", "Employee Roster Name", "Role", "Location", "Date", "Start Time", "End Time", "Break Length Minutes", "Shift Net Length", "Shift Type", "Employment Type", "Non Attended", "Status", "Published", "MealAvailabilityRequired", "MealStart", "MealEnd", "CoverageStart", "CoverageEnd", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}, MissingField.UseNull),
    SourceFile = Table.AddColumn(NeededColumns, "SourceFile", each "UNITS/" & [SourceUnit] & "/1. Input/Allocation/Published Roster.xlsx", type text),
    RenamedSourceColumns = Table.RenameColumns(SourceFile, {{"Employee_Code", "AllocationCode"}, {"Employee Code", "WorkerEmployeeID"}, {"Pay Company", "EmployerKey"}, {"Employee Roster Name", "SourceName"}, {"Role", "SourceRole"}, {"Date", "SourceDate"}, {"Shift Type", "SourceShiftType"}, {"Employment Type", "SourceEmploymentType"}, {"Break Length Minutes", "BreakMinutes"}, {"Shift Net Length", "NetRosterHours"}}),
    NormalizedText = Table.TransformColumns(RenamedSourceColumns, {{"AllocationCode", fnOvertimeRosterChecksText, type nullable text}, {"WorkerEmployeeID", fnOvertimeRosterChecksText, type nullable text}, {"EmployerKey", fnOvertimeRosterChecksText, type nullable text}, {"SourceName", fnOvertimeRosterChecksText, type nullable text}, {"SourceRole", fnOvertimeRosterChecksText, type nullable text}, {"Location", fnOvertimeRosterChecksText, type nullable text}, {"SourceEmploymentType", fnOvertimeRosterChecksText, type nullable text}, {"AgreementStream", fnOvertimeRosterChecksText, type nullable text}, {"Grade", fnOvertimeRosterChecksText, type nullable text}}),
    ParsedQuantities = Table.TransformColumns(NormalizedText, {{"BreakMinutes", fnOvertimeRosterChecksNumber, type nullable number}, {"NetRosterHours", fnOvertimeRosterChecksNumber, type nullable number}, {"SourceDate", each try Date.From(_) otherwise null, type nullable date}, {"Start Time", each try Time.From(_) otherwise null, type nullable time}, {"End Time", each try Time.From(_) otherwise null, type nullable time}}),
    ParsedFacts = Table.TransformColumns(ParsedQuantities, {{"MealAvailabilityRequired", fnOvertimeRosterChecksLogical, type nullable logical}, {"OvertimeExempt", fnOvertimeRosterChecksLogical, type nullable logical}, {"BrokenShiftException", fnOvertimeRosterChecksLogical, type nullable logical}}),
    ShiftStart = Table.AddColumn(ParsedFacts, "ShiftStart", each if [SourceDate] = null or [Start Time] = null then null else [SourceDate] & [Start Time], type nullable datetime),
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
    SourceShiftID = Table.AddColumn(DeclaredCoverage, "SourceShiftID", each fnOvertimeRosterChecksKey({[SourceUnit], [SourceFile], [SourceRow], [WorkerEmployeeID], [AllocationCode], [SourceDate], [Start Time], [End Time], [SourceRole]}), type text),
    NarrowedSource = Table.RemoveColumns(SourceShiftID, {"Start Time", "End Time", "MealStart", "MealEnd"})
in
    NarrowedSource,
OvertimeRosterChecks_WorkerPrepare = let
    RequiredColumns = {"Employee Code"},
    BDRows = fnOvertimeRosterChecksSheetPrepare(#"IMPORT BD Workers", "BD", RequiredColumns),
    TERows = fnOvertimeRosterChecksSheetPrepare(#"IMPORT TE Workers", "TE", RequiredColumns),
    JHRYRows = fnOvertimeRosterChecksSheetPrepare(#"IMPORT JH-RY Workers", "JH-RY", RequiredColumns),
    CombinedRows = Table.Combine({BDRows, TERows, JHRYRows}),
    SelectedFacts = Table.SelectColumns(CombinedRows, {"Employee Code", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Contracted FN Hours", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"}, MissingField.UseNull),
    NamedFacts = Table.RenameColumns(SelectedFacts, {{"Employee Code", "WorkerEmployeeID"}, {"Position Description", "WorkerPosition"}, {"Employee Status Description", "WorkerEmploymentType"}, {"AgreementStream", "WorkerAgreementStream"}, {"Grade", "WorkerGrade"}, {"OvertimeExempt", "WorkerOvertimeExempt"}, {"BrokenShiftException", "WorkerBrokenShiftException"}}),
    NormalizedText = Table.TransformColumns(NamedFacts, {{"WorkerEmployeeID", fnOvertimeRosterChecksText, type nullable text}, {"WorkerPosition", fnOvertimeRosterChecksText, type nullable text}, {"WorkerEmploymentType", fnOvertimeRosterChecksText, type nullable text}, {"WorkerAgreementStream", fnOvertimeRosterChecksText, type nullable text}, {"WorkerGrade", fnOvertimeRosterChecksText, type nullable text}}),
    ParsedFacts = Table.TransformColumns(NormalizedText, {{"Start Date", each try Date.From(_) otherwise null, type nullable date}, {"Termination Date", each try Date.From(_) otherwise null, type nullable date}, {"Contracted FN Hours", fnOvertimeRosterChecksNumber, type nullable number}, {"WorkerOvertimeExempt", fnOvertimeRosterChecksLogical, type nullable logical}, {"WorkerBrokenShiftException", fnOvertimeRosterChecksLogical, type nullable logical}}),
    IdentifiedWorkers = Table.SelectRows(ParsedFacts, each [WorkerEmployeeID] <> null),
    DistinctFacts = Table.Distinct(IdentifiedWorkers),
    GroupedWorkers = Table.Group(DistinctFacts, {"WorkerEmployeeID"}, {{"WorkerMatchCount", each Table.RowCount(_), Int64.Type}, {"WorkerFacts", each if Table.RowCount(_) = 1 then _{0} else [], type record}}),
    ExpandedFacts = Table.ExpandRecordColumn(GroupedWorkers, "WorkerFacts", {"WorkerPosition", "WorkerEmploymentType", "Start Date", "Termination Date", "Contracted FN Hours", "WorkerAgreementStream", "WorkerGrade", "WorkerOvertimeExempt", "WorkerBrokenShiftException"}),
    BufferedWorkers = Table.Buffer(ExpandedFacts)
in
    BufferedWorkers
in [Shifts = OvertimeRosterChecks_ShiftResults,
    Fortnights = OvertimeRosterChecks_FortnightResults,
    Segments = OvertimeRosterChecks_ExcessSegments],
Fixtures = (Evaluate as function) as table =>
let
    Anchor = #date(2026, 11, 2),
    WindowStart = #datetime(2026, 11, 2, 0, 0, 0),
    WindowEnd = #datetime(2026, 11, 16, 0, 0, 0),
    RosterColumns = {"Employee Code", "Employee_Code", "Pay Company", "Employee Roster Name", "Role", "Location", "Date", "Start Time", "End Time",
        "Break Length Minutes", "Shift Net Length", "Shift Type", "Employment Type", "MealAvailabilityRequired", "MealStart", "MealEnd",
        "CoverageStart", "CoverageEnd", "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    WorkerColumns = {"Employee Code", "Position Description", "Employee Status Description", "Start Date", "Termination Date", "Contracted FN Hours",
        "AgreementStream", "Grade", "OvertimeExempt", "BrokenShiftException"},
    AsRawSheet = (rows as list, columns as list) as table => Table.DemoteHeaders(Table.FromRecords(rows, columns, MissingField.UseNull)),
    Worker = (employee as text, optional changes as nullable record) as record => Record.Combine({
        [#"Employee Code" = employee, #"Position Description" = "AIN", #"Employee Status Description" = "Permanent Full-Time",
         #"Start Date" = #date(2020, 1, 1), #"Termination Date" = null, #"Contracted FN Hours" = 76,
         AgreementStream = "Personal Care", Grade = "1", OvertimeExempt = false, BrokenShiftException = false],
        if changes = null then [] else changes}),
    Shift = (day as number, start as time, finish as time, net as number, optional changes as nullable record) as record => Record.Combine({
        [#"Employee Code" = "E1", #"Employee_Code" = "A1", #"Pay Company" = "Synthetic employer", #"Employee Roster Name" = "Synthetic worker",
         Role = "AIN", Location = "BD", Date = Date.AddDays(Anchor, day), #"Start Time" = start, #"End Time" = finish,
         #"Break Length Minutes" = 0, #"Shift Net Length" = net, #"Shift Type" = null, #"Employment Type" = "Permanent Full-Time",
         MealAvailabilityRequired = false, MealStart = null, MealEnd = null, CoverageStart = WindowStart, CoverageEnd = WindowEnd,
         AgreementStream = "Personal Care", Grade = "1", OvertimeExempt = false, BrokenShiftException = false],
        if changes = null then [] else changes}),
    Carrier = (unit as text) as record => Shift(0, #time(1, 0, 0), #time(2, 0, 0), 1,
        [#"Employee Code" = "C" & unit, #"Employee_Code" = "AC" & unit, #"Employee Roster Name" = "Coverage carrier " & unit, Location = unit]),
    Thresholds = #table(type table [Title = text, Value = any, DateValue = any], {
        {"DailyHoursLimit", 10, null}, {"CSHospitalityHoursLimit", 8, null}, {"CSPersonalCareHoursLimit", 10, null},
        {"MinimumBreakHours", 8, null}, {"ShortGapReviewHours", 10, null}, {"MinimumFortnightDutyFreeDays", 4, null},
        {"FortnightHoursLimit", 76, null}, {"FortnightLengthDays", 14, null}, {"MinimumDutyFreeDaysBeforeNight", null, null},
        {"PayrollFortnightAnchor", null, Anchor}}),
    Periods = #table(type table [ShiftPeriod = text, StartDay = number], {{"AM", 6 / 24}, {"PM", 14 / 24}, {"NIGHT", 22 / 24}}),
    Roles = #table(type table [#"Roster Roles" = text, #"Role Group" = text], {{"AIN", "AIN"}, {"EN", "EN"}, {"RN", "RN"}, {"CS", "CS"}}),
    Facilities = #table(type table [Title = text, field_1 = text], {{"BD", "BD"}, {"TE", "TE"}, {"JH-RY", "JH-RY"}}),
    Inputs = (rows as list, optional changes as nullable record) as record =>
        let
            Base = [
                PathInput = #table({"FilePath"}, {{"C:\SyntheticResidentialCare\CLIENT\DATExx-Whiddon\2. Calculations\RosterChecks.xlsx"}}),
                PathMapping = #table({"LocalRoot", "SharePointRoot"}, {{"C:\SyntheticResidentialCare", "https://synthetic.invalid/residentialcare"}}),
                RosterBD = AsRawSheet(rows & {Carrier("BD")}, RosterColumns),
                RosterTE = AsRawSheet({Carrier("TE")}, RosterColumns),
                RosterJHRY = AsRawSheet({Carrier("JH-RY")}, RosterColumns),
                WorkersBD = AsRawSheet({Worker("E1"), Worker("CBD")}, WorkerColumns),
                WorkersTE = AsRawSheet({Worker("E1"), Worker("CTE")}, WorkerColumns),
                WorkersJHRY = AsRawSheet({Worker("E1"), Worker("CJH-RY")}, WorkerColumns),
                PeriodsBD = Periods, PeriodsTE = Periods, PeriodsJHRY = Periods,
                Thresholds = Thresholds, Roles = Roles, Facilities = Facilities],
            Result = Record.Combine({Base, if changes = null then [] else changes})
        in Result,
    ChangeThreshold = (title as text, value as any, optional dateValue as any) as table =>
        Table.FromRecords(List.Transform(Table.ToRecords(Thresholds), each if [Title] = title then
            [Title = title, Value = value, DateValue = dateValue] else _), Value.Type(Thresholds)),
    Near = (actual as any, expected as number) as logical => Value.Is(actual, type number) and Number.Abs(actual - expected) < 0.00000001,
    TargetShifts = (result as record) as table => Table.SelectRows(result[Shifts], each [EmployeeID] = "E1"),
    TargetFortnights = (result as record) as table => Table.SelectRows(result[Fortnights], each [EmployeeID] = "E1"),
    OneShift = (result as record) as record => TargetShifts(result){0},
    LaterShift = (result as record) as record => Table.SelectRows(TargetShifts(result), each [GapHours] <> null){0},
    Unknown = (value as any) as logical => Value.Is(value, type text) and Text.Contains(Text.Upper(value), "UNKNOWN"),
    Check = (name as text, input as record, predicate as function) as record =>
        let
            Attempt = try Evaluate(input),
            Assert = if Attempt[HasError] then null else try predicate(Attempt[Value]),
            Passed = if Attempt[HasError] then false else if Assert[HasError] then false else Assert[Value] = true,
            Details = if Attempt[HasError] then try Attempt[Error][Message] otherwise "Production evaluation error"
                else if Assert[HasError] then try Assert[Error][Message] otherwise "Assertion evaluation error"
                else if Passed then null else "Actual production result differs from the independent fixture expectation"
        in [Case = name, Passed = Passed, Details = Details],
    GapInputs = (laterDay as number, laterStart as time, laterFinish as time, laterNet as number, optional changes as nullable record) as record =>
        Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(laterDay, laterStart, laterFinish, laterNet)}, changes),
    DayShifts = (count as number) as list => List.Transform(List.Numbers(0, count), each Shift(_, #time(6, 0, 0), #time(14, 0, 0), 8)),
    CompleteTenDays = Inputs(DayShifts(10)),
    CompleteElevenDays = Inputs(DayShifts(11)),
    SeventySixRows = DayShifts(9) & {Shift(9, #time(6, 0, 0), #time(10, 0, 0), 4)},
    SeventyFourPlusSixRows = DayShifts(9) & {Shift(9, #time(6, 0, 0), #time(8, 0, 0), 2), Shift(10, #time(6, 0, 0), #time(12, 0, 0), 6)},
    FiftyFiveHours = Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9), Shift(2, #time(22, 0, 0), #time(6, 0, 0), 8)}),
    DailyFortnightOverlap = Inputs(DayShifts(8) & {Shift(8, #time(0, 0, 0), #time(14, 0, 0), 14)}),
    Cases = {
        Check("Daily exactly ten hours has no excess", Inputs({Shift(0, #time(6, 0, 0), #time(16, 0, 0), 10)}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("Daily eleven net hours has one excess hour", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and OneShift(r)[DailyHoursFlag] = true),
        Check("CS Hospitality exactly eight hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [Role = "CS", AgreementStream = "Hospitality"])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("CS Hospitality eight and one-quarter hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 15, 0), 8.25, [Role = "CS", AgreementStream = "Hospitality"])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0.25) and OneShift(r)[DailyHoursFlag] = true),
        Check("CS Personal Care uses ten-hour limit", Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9, [Role = "CS", AgreementStream = "Personal Care"])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0) and OneShift(r)[DailyHoursFlag] = false),
        Check("Unpaid meal is excluded once", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 10,
                [#"Break Length Minutes" = 60, MealStart = #time(12, 0, 0), MealEnd = #time(13, 0, 0)])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 0)),
        Check("Required client availability during meal counts as work", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 10,
                [#"Break Length Minutes" = 60, MealAvailabilityRequired = true, MealStart = #time(12, 0, 0), MealEnd = #time(13, 0, 0)])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1)),
        Check("Unknown meal timing preserves known daily excess quantity", Inputs({Shift(0, #time(6, 0, 0), #time(18, 0, 0), 11,
                [#"Break Length Minutes" = 60])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and Table.RowCount(Table.SelectRows(r[Segments], each Unknown([TimingStatus]))) > 0),
        Check("Seven hours fifty-nine minutes fails minimum rest", GapInputs(0, #time(21, 59, 0), #time(23, 59, 0), 2),
            (r) => Near(LaterShift(r)[GapHours], 7 + 59 / 60) and LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] = true),
        Check("Exactly eight hours passes minimum rest but needs short-gap review", GapInputs(0, #time(22, 0, 0), #time(23, 0, 0), 1),
            (r) => Near(LaterShift(r)[GapHours], 8) and LaterShift(r)[MinimumBreakFlag] = false and LaterShift(r)[ShortGapReviewFlag] = true),
        Check("Exactly ten hours has no short-gap trigger", GapInputs(1, #time(0, 0, 0), #time(2, 0, 0), 2),
            (r) => Near(LaterShift(r)[GapHours], 10) and LaterShift(r)[MinimumBreakFlag] = false and LaterShift(r)[ShortGapReviewFlag] = false),
        Check("First observed shift does not invent a predecessor", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => OneShift(r)[GapHours] = null and Unknown(OneShift(r)[GapStatus])),
        Check("Cross-unit duty participates in rest chronology", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(22, 0, 0), #time(23, 0, 0), 1, [Location = "TE"])}, RosterColumns)]),
            (r) => Table.RowCount(TargetShifts(r)) = 2 and Near(LaterShift(r)[GapHours], 8)),
        Check("Casual short rest remains a roster breach", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8, [#"Employment Type" = "Casual"]),
                Shift(0, #time(21, 59, 0), #time(23, 59, 0), 2, [#"Employment Type" = "Casual"])},
                [WorkersBD = AsRawSheet({Worker("E1", [#"Employee Status Description" = "Casual", #"Contracted FN Hours" = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Documented broken-shift exception affects overtime review only", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8),
                Shift(0, #time(21, 0, 0), #time(23, 0, 0), 2, [BrokenShiftException = true])}),
            (r) => LaterShift(r)[MinimumBreakFlag] = true and LaterShift(r)[ShortGapReviewFlag] <> true),
        Check("Four complete duty-free dates passes days-off check", CompleteTenDays,
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 4) and Near(TargetFortnights(r){0}[DaysShort], 0)),
        Check("Three complete duty-free dates flags one day short", CompleteElevenDays,
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3) and Near(TargetFortnights(r){0}[DaysShort], 1)),
        Check("Missing one unit's declared coverage leaves days off unknown", Inputs(DayShifts(10),
                [RosterTE = AsRawSheet({Record.Combine({Carrier("TE"), [CoverageStart = null, CoverageEnd = null]})}, RosterColumns)]),
            (r) => Unknown(TargetFortnights(r){0}[DaysOffStatus]) and TargetFortnights(r){0}[DutyFreeDays] = null),
        Check("Overnight duty occupies both touched calendar dates", Inputs({Shift(0, #time(22, 0, 0), #time(6, 0, 0), 8)} &
                List.Transform({2..10}, each Shift(_, #time(6, 0, 0), #time(7, 0, 0), 1))),
            (r) => Near(TargetFortnights(r){0}[DutyFreeDays], 3)),
        Check("Monday fifteen to Wednesday twenty-two gives fifty-five hours one free day", FiftyFiveHours,
            (r) => Near(LaterShift(r)[GapHours], 55) and Near(LaterShift(r)[NightDutyFreeDays], 1) and LaterShift(r)[NightTransitionFlag] = true),
        Check("Blank before-night threshold does not invent a pass or fail", FiftyFiveHours,
            (r) => LaterShift(r)[NightDaysFlag] = null and Near(LaterShift(r)[NightDutyFreeDays], 1)),
        Check("Consecutive night shifts do not create a non-night transition", Inputs({Shift(0, #time(22, 0, 0), #time(6, 0, 0), 8), Shift(1, #time(22, 0, 0), #time(6, 0, 0), 8)}),
            (r) => LaterShift(r)[NightTransitionFlag] = false),
        Check("Exactly seventy-six net hours has no fortnight excess", Inputs(SeventySixRows),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 76) and Near(TargetFortnights(r){0}[FortnightExcessHours], 0)),
        Check("Seventy-four plus six identifies four excess hours", Inputs(SeventyFourPlusSixRows),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 80) and Near(TargetFortnights(r){0}[FortnightExcessHours], 4)
                and Near(List.Sum(TargetShifts(r)[FortnightExcessHours]), 4)),
        Check("Daily overtime hours still enter the net fortnight counter once", Inputs(List.Transform({0..5}, each Shift(_, #time(0, 0, 0), #time(10, 0, 0), 10)) &
                {Shift(6, #time(0, 0, 0), #time(14, 0, 0), 14), Shift(7, #time(0, 0, 0), #time(6, 0, 0), 6)}),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 80) and Near(TargetFortnights(r){0}[FortnightExcessHours], 4)),
        Check("Daily and fortnight overlapping excess is counted once", DailyFortnightOverlap,
            (r) => Near(TargetFortnights(r){0}[FortnightExcessHours], 2) and Near(List.Sum(r[Segments][SegmentHours]), 4)
                and Table.RowCount(Table.SelectRows(r[Segments], each [DailyTrigger] and [FortnightTrigger])) > 0),
        Check("A known boundary-crossing shift allocates two and six hours to adjacent fortnights", Inputs({Shift(13, #time(22, 0, 0), #time(6, 0, 0), 8,
                [CoverageEnd = #datetime(2026, 11, 30, 0, 0, 0)])}, [
                RosterTE = AsRawSheet({Record.Combine({Carrier("TE"), [CoverageEnd = #datetime(2026, 11, 30, 0, 0, 0)]})}, RosterColumns),
                RosterJHRY = AsRawSheet({Record.Combine({Carrier("JH-RY"), [CoverageEnd = #datetime(2026, 11, 30, 0, 0, 0)]})}, RosterColumns)]),
            (r) => List.Sort(TargetFortnights(r)[NetRosterHours]) = {2, 6}),
        Check("Unknown meal placement across a fortnight boundary is not prorated", Inputs({Shift(13, #time(20, 0, 0), #time(8, 0, 0), 11,
                [#"Break Length Minutes" = 60, CoverageEnd = #datetime(2026, 11, 30, 0, 0, 0)])}),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and List.AllTrue(List.Transform(TargetFortnights(r)[FortnightStatus], Unknown))),
        Check("Exact duplicate source rows never double known net hours", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8), Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)}),
            (r) => Near(TargetFortnights(r){0}[NetRosterHours], 8) or Unknown(TargetFortnights(r){0}[FortnightStatus])),
        Check("Overlapping cross-unit duties are explicitly unknown", Inputs({Shift(0, #time(6, 0, 0), #time(14, 0, 0), 8)},
                [RosterTE = AsRawSheet({Carrier("TE"), Shift(0, #time(10, 0, 0), #time(18, 0, 0), 8, [Location = "TE"])}, RosterColumns)]),
            (r) => List.AllTrue(List.Transform(TargetShifts(r)[OverlapStatus], Unknown))),
        Check("Conflicting worker rows cannot silently choose employment type", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [WorkersBD = AsRawSheet({Worker("E1"), Worker("E1", [#"Employee Status Description" = "Casual"]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[IdentityStatus]) or Unknown(OneShift(r)[DailyStatus])),
        Check("Missing CS agreement stream remains unknown", Inputs({Shift(0, #time(6, 0, 0), #time(15, 0, 0), 9, [Role = "CS", AgreementStream = null])},
                [WorkersBD = AsRawSheet({Worker("E1", [#"Position Description" = "CS", AgreementStream = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[DailyStatus]) and OneShift(r)[DailyHoursFlag] = null),
        Check("Invalid daily threshold does not default or disable gap measurement", GapInputs(0, #time(21, 59, 0), #time(23, 59, 0), 2,
                [Thresholds = ChangeThreshold("DailyHoursLimit", "ten")]),
            (r) => Unknown(LaterShift(r)[DailyStatus]) and LaterShift(r)[MinimumBreakFlag] = true),
        Check("Duplicate threshold title is not arbitrarily selected", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [Thresholds = Table.Combine({Thresholds, #table(type table [Title = text, Value = any, DateValue = any], {{"DailyHoursLimit", 12, null}})})]),
            (r) => Unknown(OneShift(r)[DailyStatus])),
        Check("Missing payroll anchor leaves fortnight unknown without hiding daily excess", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11)},
                [Thresholds = ChangeThreshold("PayrollFortnightAnchor", null, null)]),
            (r) => Near(OneShift(r)[DailyExcessHours], 1) and Unknown(OneShift(r)[FortnightStatus])),
        Check("Missing payroll eligibility remains explicit", Inputs({Shift(0, #time(6, 0, 0), #time(17, 0, 0), 11,
                [OvertimeExempt = null, Grade = null, AgreementStream = null])},
                [WorkersBD = AsRawSheet({Worker("E1", [OvertimeExempt = null, Grade = null, AgreementStream = null]), Worker("CBD")}, WorkerColumns)]),
            (r) => Unknown(OneShift(r)[DailyStatus]) or Unknown(OneShift(r)[OvertimeStatus]))
    },
    Results = Table.FromRecords(Cases, type table [Case = text, Passed = logical, Details = nullable text])
in
    Results
in Fixtures(Evaluate);
