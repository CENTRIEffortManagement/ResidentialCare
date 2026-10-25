section BDBalanceRemovalFixtures;

// Query: BDBalanceRemovalFixtures
// Purpose: Exercise actual BD production expressions with isolated synthetic worker sources.
shared BDBalanceRemovalFixtures = (Evaluate as function, EvaluateStaff as function) as table => let
    Start = #date(2026,10,1),
    Facilities = #table({"Title","FacilityShort","Facility-Abbrev"},{{"BD: Beaudesert","BD","BD"},{"TE: Temora","TE","TE"}}),
    RosterColumns = {"Location","Department","Employee Roster Name","Employment Type","Role","Employee Code","Date"},
    AvailabilityColumns = {"Site Name","Employee Name","PayrollCode","Department","DateTimeStart","DateTimeEnd","Comments","Type"},
    HRColumns = {"Employee Code","Location Description","Position Description","Employee Status Description","Start Date","Termination Date","Contracted FN Hours"},
    RawSheet = (records as list, columns as list) as table => Table.DemoteHeaders(Table.FromRecords(records,columns,MissingField.UseNull)),
    Roster = (id as text, optional role as nullable text, optional site as nullable text, optional name as nullable text, optional employment as nullable text) as record =>
        [Location=site ?? "BD: Beaudesert", Department="Care", #"Employee Roster Name"=name ?? ("Worker " & id),
            #"Employment Type"=employment ?? "Permanent Part Time", Role=role ?? "Assistant in Nursing AM", #"Employee Code"=id, Date=Start],
    RecordRow = (id as text, optional reason as nullable text, optional site as nullable text, optional name as nullable text) as record =>
        [#"Site Name"=site ?? "BD: Beaudesert", #"Employee Name"=name ?? ("Worker " & id), PayrollCode=id, Department="Care",
            DateTimeStart=#datetime(2026,10,1,0,0,0), DateTimeEnd=#datetime(2026,10,2,0,0,0), Comments=reason ?? "AVAIL", Type=reason ?? "AVAIL"],
    HR = (id as text, optional role as nullable text, optional site as nullable text, optional termination as nullable date, optional hours as nullable number) as record =>
        [#"Employee Code"=id, #"Location Description"=site ?? "BD: Beaudesert", #"Position Description"=role ?? "Assistant in Nursing",
            #"Employee Status Description"="Permanent Part Time", #"Start Date"=#date(2020,1,1), #"Termination Date"=termination, #"Contracted FN Hours"=hours ?? 76],
    Inputs = (roster as list, availability as list, hr as list, optional unit as nullable text) as record =>
        [Roster=RawSheet({[Location="BD: Beaudesert",Department="Care",#"Employee Roster Name"=null,#"Employment Type"="Agency",Role=null,#"Employee Code"=null,Date=Start]} & roster,RosterColumns),
            Availability=RawSheet(availability,AvailabilityColumns),HR=RawSheet(hr,HRColumns),
            Facilities=if unit="BD-FULL" then #table({"Title","FacilityShort","Facility-Abbrev"},{{"BD: Beaudesert","BD","BD-FULL"}}) else Facilities,Unit=unit ?? "BD"],
    Check = (name as text, run as function, predicate as function) as record => let
        Attempt=try run(), Verdict=if Attempt[HasError] then [HasError=false,Value=false] else try predicate(Attempt[Value]),
        Passed=not Attempt[HasError] and not Verdict[HasError] and Verdict[Value]=true,
        Details=if Attempt[HasError] then (try Attempt[Error][Message] otherwise "Evaluation failed")
            else if Verdict[HasError] then (try Verdict[Error][Message] otherwise "Assertion failed") else if Passed then null else "Unexpected result"
    in [Case=name,Passed=Passed,Details=Details],
    Membership = (inputs as record) as table => Table.Buffer(Evaluate(inputs)[Membership]),
    ErrorBlocked = (inputs as record) as logical => (try Table.RowCount(Membership(inputs)))[HasError],
    Cases = {
        Check("roster-only worker",()=>Membership(Inputs({Roster("1001")},{},{HR("1001")})),each Table.RowCount(_)=1 and _{0}[In Allocation]=true and _{0}[In Records]=false and _{0}[EmployeeID]="1001"),
        Check("availability-only worker",()=>Membership(Inputs({},{RecordRow("1002")},{HR("1002")})),each Table.RowCount(_)=1 and _{0}[In Allocation]=false and _{0}[In Records]=true),
        Check("both sources one worker",()=>Membership(Inputs({Roster("1003")},{RecordRow("1003")},{HR("1003")})),each Table.RowCount(_)=1 and _{0}[In Allocation]=true and _{0}[In Records]=true),
        Check("HR-only and hypothetical Balance-only IDs excluded",()=>Membership(Inputs({Roster("1001")},{},{HR("1001"),HR("HR-ONLY"),HR("BALANCE-ONLY")})),each Table.RowCount(_)=1 and _{0}[EmployeeID]="1001"),
        Check("wholly UNAVAIL worker retained",()=>Evaluate(Inputs({},{RecordRow("1004","UNAVAIL")},{HR("1004")}))[Eligible],each Table.RowCount(_)=1 and _{0}[EmployeeID]="1004"),
        Check("roster facility retained against HR home",()=>Membership(Inputs({Roster("1005")},{},{HR("1005",null,"TE: Temora")})),each Table.RowCount(_)=1 and _{0}[#"Facility-Abbrev"]="BD"),
        Check("availability wrong HR home excluded",()=>Membership(Inputs({},{RecordRow("wrong-home")},{HR("wrong-home",null,"TE: Temora")})),each Table.IsEmpty(_)),
        Check("source alias maps complete canonical facility key",()=>Evaluate(Inputs({Roster("1006")},{RecordRow("1006")},{HR("1006")},"BD-FULL"))[Prepared],each Table.RowCount(_)=1 and _{0}[#"Facility-Abbrev"]="BD-FULL"),
        Check("complete text employee ID preserved",()=>Membership(Inputs({Roster("001234-X")},{RecordRow("001234-X")},{HR("001234-X")})),each Table.RowCount(_)=1 and _{0}[EmployeeID]="001234-X"),
        Check("agency roster excluded",()=>Membership(Inputs({Roster("AGENCY",null,null,null,"Agency ")},{},{})),each Table.IsEmpty(_)),
        Check("HF and extraction-error records excluded",()=>Membership(Inputs({},{RecordRow("HF"),RecordRow("999999")},{HR("HF"),HR("999999")})),each Table.IsEmpty(_)),
        Check("termination before roster start excluded",()=>Membership(Inputs({Roster("1007")},{},{HR("1007",null,null,#date(2026,9,30))})),each Table.IsEmpty(_)),
        Check("termination on roster start retained",()=>Membership(Inputs({Roster("1008")},{},{HR("1008",null,null,Start)})),each Table.RowCount(_)=1),
        Check("conflicting eligible HR records block",()=>ErrorBlocked(Inputs({Roster("1009")},{},{HR("1009"),HR("1009","Enrolled Nurse")})),each _=true),
        Check("identical duplicate HR rows allowed",()=>Membership(Inputs({Roster("1010")},{},{HR("1010"),HR("1010")})),each Table.RowCount(_)=1),
        Check("preferred roster role aligned with HR",()=>Membership(Inputs({Roster("1011","Assistant in Nursing AM"),Roster("1011","Enrolled Nurse AM")},{},{HR("1011","Enrolled Nurse")})),each Table.RowCount(_)=1 and _{0}[Role]="Enrolled Nurse"),
        Check("roster role before HR fallback",()=>Membership(Inputs({Roster("1012","Enrolled Nurse AM")},{},{HR("1012","Assistant in Nursing")})),each Table.RowCount(_)=1 and _{0}[Role]="Enrolled Nurse"),
        Check("availability role from HR",()=>Membership(Inputs({},{RecordRow("1013")},{HR("1013","Enrolled Nurse")})),each Table.RowCount(_)=1 and _{0}[Role]="Enrolled Nurse"),
        Check("availability with missing HR excluded",()=>Membership(Inputs({},{RecordRow("missing-HR")},{})),each Table.IsEmpty(_)),
        Check("no Balance role fallback",()=>ErrorBlocked(Inputs({Record.TransformFields(Roster("1014"),{{"Role",each null}})},{},{})),each _=true),
        Check("leave provenance compatibility false",()=>Membership(Inputs({Roster("1015")},{RecordRow("1015")},{HR("1015")})),each Table.RowCount(_)=1 and List.AllTrue(List.Transform(_[In Leave Balances],each _=false))),
        Check("production eligibility and checks",()=>Evaluate(Inputs({Roster("1016")},{RecordRow("1016")},{HR("1016")})),each Table.RowCount([Eligible])=1 and List.AllTrue(List.Transform([Checks][Status],each _="Pass"))),
        Check("reported contract retained",()=>Evaluate(Inputs({Roster("1017")},{},{HR("1017",null,null,null,57)}))[Eligible],each Table.RowCount(_)=1 and _{0}[Contracted FN Hours]=57),
        Check("roster without HR retains null contract",()=>Evaluate(Inputs({Roster("no-HR")},{},{}))[Eligible],each Table.RowCount(_)=1 and _{0}[Contracted FN Hours]=null)
    },
    StaffRow = (id as text, hours as nullable number, cap as nullable number, applied as nullable number, issue as nullable text) as record =>
        [Name="Staff "&id,Role="AIN",Resource=1,Misalignment=0,Source="Availability",EmployeeID=id,#"Facility-Abbrev"="BD",EmploymentType=if hours=null then "Casual" else "Permanent Part Time",PreferredRole="AIN",
            #"In Allocation"=true,#"In Records"=true,#"Worker Membership"="Both",#"Worker Record Status"="Found",#"Worker Contract Issue"=null,#"Contracted FN Hours"=hours,
            RosterShiftCap=cap,AppliedCasualRosterShiftCap=applied,CasualCapIssue=issue],
    StaffCases={
        Check("staff reported contract proration and cap unchanged",()=>EvaluateStaff([Staff=Table.FromRecords({StaffRow("contract",76,null,null,null)})])[Published]{0},each Number.Abs([Contracted Roster Hours]-76*29/14)<0.00000001 and [Contracted Shifts]=20 and [Effective Shift Cap]=20 and [Contract Source]="Worker record"),
        Check("staff casual roster cap unchanged",()=>EvaluateStaff([Staff=Table.FromRecords({StaffRow("casual",null,13,13,null)})])[Published]{0},each [Effective Shift Cap]=13 and [Contracted FN Hours]=null and [Contracted Roster Hours]=null),
        Check("staff casual Settings fallback unchanged",()=>EvaluateStaff([Staff=Table.FromRecords({StaffRow("fallback",null,null,40,"Missing roster cap")})])[Published]{0},each [Effective Shift Cap]=40 and [Limit Basis]="Casual fallback - Settings maximum")
    }
in Table.FromRecords(Cases & StaffCases);
