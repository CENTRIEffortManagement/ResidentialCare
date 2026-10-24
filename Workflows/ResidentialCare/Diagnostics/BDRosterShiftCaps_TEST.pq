section BDRosterShiftCapsFixtures;

// Query: BDRosterShiftCapsFixtures
// Purpose: Exercise the actual BD roster-cap, availability and resource-contract queries with synthetic workers.
// Inputs: Evaluators are assembled from approved M source by scripts/test-bd-roster-shift-caps-native.ps1.
// Output: One assertion per ID join, precedence, fallback, role-share or publication scenario.
// Notes: Fixtures have no workbook, file or network access and do not refresh production inputs.
// Notes: Synthetic caps replace the typed IMPORT output; defensive fractional cases do not test its retained Int64 coercion.
shared BDRosterShiftCapsFixtures = (Availability as function, StaffList as function) as table =>
let
    Near = (A as number, B as number) as logical => Number.Abs(A - B) < 0.00000001,
    Check = (Name as text, Work as function) as record =>
        let Result = try Work()
        in [Check = Name, Passed = if Result[HasError] then false else Result[Value] = true,
            Details = if Result[HasError] then (try Result[Error][Message] otherwise "Evaluation failed") else null],
    Change = (Base as record, Overrides as record) as record => Record.Combine({Base, Overrides}),
    Workers = #table(type table [EmployeeID = text, #"Facility-Abbrev" = text, Name = text,
        Role = text, PreferredRole = text, EmploymentType = text, #"Contracted FN Hours" = nullable number,
        #"In Allocation" = logical, #"In Records" = logical, #"Worker Membership" = text], {
        {"001", "BD", "Contracted worker", "AIN", "AIN", "Casual", 38, true, true, "Both"},
        {"002", "BD", "Roster worker", "AIN", "AIN", "Permanent", null, true, true, "Both"},
        {"003", "BD", "Maximum worker", "AIN", "AIN", "Casual", 0, false, true, "Records only"},
        {"004", "BD", "Duplicate worker", "AIN", "AIN", "Permanent", 0, true, true, "Both"},
        {"005", "BD", "Missing worker", "AIN", "AIN", "Casual", null, true, false, "Allocation only"},
        {"006", "BD", "Zero worker", "AIN", "AIN", "Permanent", 0, true, true, "Both"},
        {"007", "BD", "Negative worker", "AIN", "AIN", "Casual", null, true, true, "Both"},
        {"008", "BD", "Error worker", "AIN", "AIN", "Casual", null, true, true, "Both"},
        {"009", "BD", "Fractional worker", "AIN", "AIN", "Casual", null, true, true, "Both"},
        {"010", "BD", "Blank worker", "AIN", "AIN", "Casual", null, true, true, "Both"},
        {"011", "BD", "Conflicting worker", "AIN", "AIN", "Casual", null, true, true, "Both"},
        {"012", "BD", "Text worker", "AIN", "AIN", "Casual", null, true, true, "Both"},
        {"013", "BD", "NaN worker", "AIN", "AIN", "Casual", null, true, true, "Both"},
        {"014", "BD", "Infinite worker", "AIN", "AIN", "Casual", null, true, true, "Both"}
    }),
    Caps = #table(type table [#"Employee Code" = any, RosterShiftCap = any, Unit = text], {
        {"001", 3, "BD"}, {" 002 ", 8, "Another analysis facility"}, {"003", 30, "BD"},
        {"004", 10, "BD"}, {"004", 10, "Another analysis facility"},
        {"006", 0, "BD"}, {"007", -1, "BD"}, {"008", error "Invalid imported cap cell", "BD"},
        {"009", 2.5, "BD"}, {"010", null, "BD"}, {"011", 5, "BD"}, {"011", 6, "BD"},
        {"012", "not a cap", "BD"}, {"013", #nan, "BD"}, {"014", #infinity, "BD"}
    }),
    MakeShifts = (ID as text, Name as text, Role as text, Count as number) as table =>
        #table(type table [ID = text, #"Facility-Abbrev" = text, Role = text, Name = text],
            List.Transform({1..Count}, each {ID, "BD", Role, Name})),
    DayShifts = Table.Combine({MakeShifts("001", "Contracted worker", "AIN", 28),
        MakeShifts("002", "Roster worker", "AIN", 28), MakeShifts("002", "Roster worker", "RN", 28),
        MakeShifts("003", "Maximum worker", "AIN", 10), MakeShifts("004", "Duplicate worker", "AIN", 4)}),
    Shares = #table(type table [EmployeeID = text, #"Facility-Abbrev" = text,
        Role = text, RoleCount = Int64.Type, RoleShare = number], {
        {"002", "BD", "AIN", 2, 0.75}, {"002", "BD", "RN", 2, 0.25}
    }),
    Inputs = [Staff = Workers, Caps = Caps, Maximum = 20, DayShifts = DayShifts, Shares = Shares],
    Result = Availability(Inputs),
    Published = Result[Staff],
    Lookup = Result[Lookup],
    Exceptions = Result[Exceptions],
    ByID = (Rows as table, ID as text) as record => Table.SelectRows(Rows, each Text.Trim([EmployeeID]) = ID){0},
    ByRole = (Rows as table, Name as text, Role as text) as record =>
        Table.SelectRows(Rows, each [Name] = Name and [Role] = Role){0},
    MasterInput = (Staff as table) as table =>
        let
            Resources = Table.AddIndexColumn(Staff, "Resource", 1, 1, Int64.Type),
            WithSource = Table.AddColumn(Resources, "Source", each "Availability", type text),
            WithMisalignment = Table.AddColumn(WithSource, "Misalignment", each null, type nullable text),
            WithRecordStatus = Table.AddColumn(WithMisalignment, "Worker Record Status", each "Matched", type text),
            WithContractIssue = Table.AddColumn(WithRecordStatus, "Worker Contract Issue", each null, type nullable text)
        in WithContractIssue,
    ContractInputs = [Staff = MasterInput(Published), Maximum = 20],
    ContractResult = StaffList(ContractInputs),
    Contracts = ContractResult[Contracts],
    Master = ContractResult[Master],
    NoHoursIDs = {"002", "003", "004", "005", "006", "007", "008", "009", "010", "011", "012", "013", "014"},
    ExceptionIDs = {"005", "006", "007", "008", "009", "010", "011", "012", "013", "014"},
    ExpectedMasterColumns = {"Name", "Role", "Resource", "Misalignment", "Source", "EmployeeID", "Facility-Abbrev",
        "EmploymentType", "PreferredRole", "In Allocation", "In Records", "Worker Membership", "Worker Record Status",
        "Worker Contract Issue", "Contract Source", "Contracted FN Hours", "Roster Start", "Roster End", "Roster Fortnights",
        "Contracted Roster Hours", "Contracted Shift Equivalent", "Contracted Shifts", "Settings Max Availability",
        "Effective Shift Cap", "Limit Basis"},
    UpdateHours = (Hours as any) as table =>
        Table.TransformColumns(Table.FirstN(ContractInputs[Staff], 1), {{"Contracted FN Hours", each Hours, type any}}),
    RejectInvalidContract = (Hours as any) as logical =>
        let Attempt = try Table.RowCount(StaffList(Change(ContractInputs, [Staff = UpdateHours(Hours)]))[Master])
        in Attempt[HasError],
    RejectCapSchema = (Column as text) as logical =>
        let Attempt = try Table.RowCount(Availability(Change(Inputs, [Caps = Table.RemoveColumns(Caps, {Column})]))[Lookup])
        in Attempt[HasError],
    Rows = {
        Check("ID join trims cap identifiers and preserves leading zeroes", () =>
            ByID(Published, "002")[RosterShiftCap] = 8 and ByID(Published, "001")[RosterShiftCap] = 3),
        Check("Lookup matches by ID regardless of the source facility label", () =>
            ByID(Published, "002")[AppliedCasualRosterShiftCap] = 8),
        Check("Positive hours override a historical cap regardless of employment label", () =>
            ByID(Published, "001")[AppliedCasualRosterShiftCap] = null
            and ByID(Contracts, "001")[#"Contracted FN Hours"] = 38
            and Near(ByID(Contracts, "001")[#"Effective Shift Cap"], 10)
            and ByID(Contracts, "001")[#"Contract Source"] = "Worker record"),
        Check("Blank hours use the joined whole-roster cap for a permanent label", () =>
            ByID(Published, "002")[AppliedCasualRosterShiftCap] = 8
            and ByID(Contracts, "002")[#"Effective Shift Cap"] = 8),
        Check("Zero hours use a cap bounded by the Settings maximum", () =>
            ByID(Published, "003")[RosterShiftCap] = 30
            and ByID(Published, "003")[AppliedCasualRosterShiftCap] = 20),
        Check("Identical duplicate caps resolve once without worker fan-out", () =>
            Table.RowCount(Table.SelectRows(Lookup, each [EmployeeID] = "004")) = 1
            and ByID(Published, "004")[AppliedCasualRosterShiftCap] = 10
            and Table.RowCount(Published) = Table.RowCount(Workers)),
        Check("Missing zero negative error fractional blank text nonfinite and conflicting caps fall back", () =>
            List.AllTrue(List.Transform(ExceptionIDs, each
                ByID(Published, _)[AppliedCasualRosterShiftCap] = 20
                and ByID(Published, _)[CasualCapIssue] <> null))),
        Check("Exceptions have exactly the affected employee IDs and no duplicate workers", () =>
            List.Sort(Exceptions[EmployeeID]) = List.Sort(ExceptionIDs)
            and Table.RowCount(Exceptions) = List.Count(ExceptionIDs)),
        Check("Exception diagnostics expose reported text without propagating imported cell errors", () =>
            Table.ColumnNames(Exceptions) = {"EmployeeID", "Facility-Abbrev", "Name", "ReportedRosterShiftCap",
                "AppliedCasualRosterShiftCap", "CasualCapIssue"}
            and List.AllTrue(List.Transform(Table.ToRecords(Exceptions), each
                let TextAttempt = try [ReportedRosterShiftCap]
                in not TextAttempt[HasError] and [AppliedCasualRosterShiftCap] = 20 and [CasualCapIssue] <> null))),
        Check("Conflicting duplicate caps do not choose an arbitrary cap", () =>
            ByID(Published, "011")[RosterShiftCap] = null and ByID(Published, "011")[AppliedCasualRosterShiftCap] = 20),
        Check("An empty lookup falls back for workers without hours only", () =>
            let Empty = Availability(Change(Inputs, [Caps = Table.FirstN(Caps, 0)]))
            in Table.RowCount(Empty[Exceptions]) = List.Count(NoHoursIDs)
                and ByID(Empty[Staff], "001")[AppliedCasualRosterShiftCap] = null),
        Check("A source ID shared across facilities is resolved once by ID", () =>
            ByID(Lookup, "004")[RosterShiftCap] = 10 and ByID(Lookup, "004")[CapLookupIssue] = null),
        Check("Role shares apportion counts above the resolved whole-roster cap", () =>
            Near(ByRole(Result[RoleResources], "Roster worker", "AIN")[AvailabilityCapped], 6)
            and Near(ByRole(Result[RoleResources], "Roster worker", "RN")[AvailabilityCapped], 2)
            and ByRole(Result[RoleResources], "Roster worker", "AIN")[StdRosterDays] = 8),
        Check("Counts below the resolved cap retain the existing unscaled arithmetic", () =>
            ByRole(Result[RoleResources], "Maximum worker", "AIN")[AvailabilityCapped] = 10
            and ByRole(Result[RoleResources], "Duplicate worker", "AIN")[AvailabilityCapped] = 4),
        Check("Positive-contract availability retains the existing Settings cap", () =>
            ByRole(Result[RoleResources], "Contracted worker", "AIN")[StdRosterDays] = 20
            and ByRole(Result[RoleResources], "Contracted worker", "AIN")[AvailabilityCapped] = 20),
        Check("Role output still sums the published capped worker-role values", () =>
            Table.SelectRows(Result[Roles], each [Role] = "AIN"){0}[AvailabilityCapped] = 40
            and Table.SelectRows(Result[Roles], each [Role] = "RN"){0}[AvailabilityCapped] = 2),
        Check("No-hours contract quantities remain N/A rather than fabricated hours", () =>
            List.AllTrue(List.Transform(NoHoursIDs, each
                let Contract = ByID(Contracts, _)
                in Contract[#"Contracted FN Hours"] = null and Contract[#"Contracted Roster Hours"] = null
                    and Contract[#"Contracted Shift Equivalent"] = null and Contract[#"Contracted Shifts"] = null
                    and Contract[#"Contract Source"] = "Casual roster cap"))),
        Check("Availability and resource effective caps agree for every no-hours worker", () =>
            List.AllTrue(List.Transform(NoHoursIDs, each
                ByID(Published, _)[AppliedCasualRosterShiftCap] = ByID(Contracts, _)[#"Effective Shift Cap"]))),
        Check("Limit Basis distinguishes roster caps from Settings fallbacks", () =>
            ByID(Contracts, "002")[#"Limit Basis"] = "Casual fallback - RosterShiftCap"
            and ByID(Contracts, "005")[#"Limit Basis"] = "Casual fallback - Settings maximum"),
        Check("Published Masterlist retains its existing columns and worker count", () =>
            Table.ColumnNames(Master) = ExpectedMasterColumns and Table.RowCount(Master) = Table.RowCount(Workers)),
        Check("Valid fallback exceptions do not block resource publication", () =>
            Table.IsEmpty(Table.SelectRows(ContractResult[Checks], each [Status] = "Fail"))),
        Check("The Settings maximum remains configurable", () =>
            let ChangedAvailability = Availability(Change(Inputs, [Maximum = 12])),
                ChangedContracts = StaffList([Staff = MasterInput(ChangedAvailability[Staff]), Maximum = 12])
            in ByID(ChangedAvailability[Staff], "003")[AppliedCasualRosterShiftCap] = 12
                and ByID(ChangedAvailability[Staff], "005")[AppliedCasualRosterShiftCap] = 12
                and ByID(ChangedContracts[Contracts], "003")[#"Effective Shift Cap"] = 12),
        Check("Invalid existing contracted hours remain blocking", () =>
            List.AllTrue(List.Transform({-1, #nan, #infinity, "invalid hours"}, RejectInvalidContract))),
        Check("Missing required cap columns remain source schema errors", () =>
            RejectCapSchema("Employee Code") and RejectCapSchema("RosterShiftCap"))
    }
in
    Table.FromRecords(Rows, type table [Check = text, Passed = logical, Details = nullable text]);
