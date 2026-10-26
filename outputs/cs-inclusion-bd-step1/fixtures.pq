section CSRoleDemandFixtures;

// Query: CSRoleDemandFixtures
// Purpose: Exercise actual generic demand handover expressions with synthetic boundary tables.
// Inputs: Evaluate is assembled from the exact edited organisation Manual Read source.
// Output: Independently asserted pass/fail records; no external sources or workbook refresh.
shared CSRoleDemandFixtures = (Evaluate as function) as table => let
    Near = (A as number, B as number) as logical => Number.Abs(A-B) <= 0.0000001,
    Change = (Base as record, Fields as record) as record => Record.Combine({Base,Fields}),
    Set = (Row as record, Field as text, Value as any) as record => Record.TransformFields(Row,{{Field,each Value}}),
    UpdateFirst = (T as table, Field as text, Value as any) as table =>
        Table.FromRecords({Set(T{0},Field,Value)} & List.Skip(Table.ToRecords(T))),
    Check = (Name as text, Work as function) as record => let
        Attempt = try Work(),
        Passed = if Attempt[HasError] then false else Attempt[Value] = true
    in [Case=Name,Passed=Passed,Details=if Attempt[HasError] then Attempt[Error][Message] else if Passed then null else "Unexpected result"],
    ErrorChecks = (T as table) as table => Table.SelectRows(T,each
        Record.FieldOrDefault(_,"Severity",null) = "Error"
        or List.Contains({"Fail","Error","FAIL","ERROR"},Record.FieldOrDefault(_,"Status",null))),
    Reject = (I as record) as logical => let
        R = Evaluate(I), Checks = try Table.Buffer(R[Checks]),
        Publication = try Table.RowCount(Table.Buffer(R[Demand]))
    in not Checks[HasError] and Table.RowCount(ErrorChecks(Checks[Value])) > 0 and Publication[HasError],
    Roles = #table({"Roster Roles","DC Role","Role Group","DC Category","Direct Care %"},{
        {"Registered Nurse","RN","RN","RN",1},
        {"Assistant in Nursing","AIN","AIN","OTHER",1},
        {"Enrolled Nurse","EN","EN","OTHER",0.5},
        {"Care Services Employee",null,"CS","NA",null},
        {"Care Services Employee alternate",null,"CS","NA",null},
        {"Care Services Hospitality",null,"CSH","NA",null}
    }),
    Analysis = #table({"Roles","Effort Management Analysis","Leave Balance Analysis"},{
        {"RN",true,true},{"AIN",true,true},{"EN",true,true},{"CS",true,true},{"CSH",true,true}
    }),
    Units = #table({"Location","Unit"},{{"BD: Beaudesert","BD"},{"TE: Temora","TE"}}),
    UnitAnalysis = #table({"Title","Effort Management Analysis"},{{"BD",true},{"TE",false}}),
    ShiftEnds = #table(type table [Shift=text,End=time],{{"AM",#time(14,0,0)},{"PM",#time(21,0,0)},{"NS",#time(6,0,0)}}),
    Days = {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"},
    RawColumns = {"Location","Role","Employee Code","Week No","Week Day","Start Time","End Time","Shift Net Length","Roster Hours"},
    CSAmount = (W as number, D as number, S as number) as number => (if W=1 then 5 else 8) + D/10 + S/100,
    MakeRaw = (Label as text, Hours as function) as table => Table.FromRecords(List.Combine(List.Transform({1..2},(W) =>
        List.Combine(List.Transform({1..7},(D) => List.Transform({1..3},(S) =>
            [Location="BD: Beaudesert",Role=Label,#"Employee Code"=100+S,#"Week No"=29+W,
                #"Week Day"=Days{D-1},#"Start Time"={#time(7,0,0),#time(15,0,0),#time(22,0,0)}{S-1},
                #"End Time"={#time(14,0,0),#time(21,0,0),#time(6,0,0)}{S-1},
                #"Shift Net Length"=Hours(W,D,S),#"Roster Hours"=9999]
        ))))),RawColumns),
    CSRaw = MakeRaw("Care Services Employee",CSAmount),
    CareRaw = MakeRaw("Registered Nurse",(W,D,S) => 1),
    HospitalityRaw = MakeRaw("Care Services Hospitality",(W,D,S) => 1000),
    Master = Table.Combine({CSRaw,CareRaw,HospitalityRaw}),
    CareDefinitions = {[Role="RN",Category="RN",Percentage=1,Base=8],
        [Role="AIN",Category="OTHER",Percentage=1,Base=10],[Role="EN",Category="OTHER",Percentage=0.5,Base=4]},
    CareRows = Table.FromRecords(List.Combine(List.Transform(CareDefinitions,(R) =>
        List.Combine(List.Transform({1..14},(D) => List.Transform({1..3},(S) => let
            W = 1 + Number.IntegerDivide(D-1,7), Day = Number.Mod(D-1,7)+1,
            H = R[Base] + W + Day/10 + S/100
        in [Facility="BD",Role=R[Role],#"DC Role"=R[Role],#"DC Category"=R[Category],
            MinuteCategory=if R[Category]="RN" then "RN" else "OTHERS",#"Direct Care %"=R[Percentage],
            #"Week No"=29+W,FortnightWeek=W,DayOfWeek=Day,#"Week Day"=Days{Day-1},
            FortnightDayIndex=D,FortnightDay=Text.From(W)&"-"&Days{Day-1},Shift={"AM","PM","NS"}{S-1},ShiftIndex=S,
            RedistributedDemandHours=H,FTE=H/7.6,WeekdayShiftTargetMinutes=H*R[Percentage]*60,
            HistoricalRosterHours=H/1.25,HistoricalProductiveHours=H/1.25*R[Percentage],
            HistoricalCoverageStatus="PASS",ProfileAlignmentStatus="PASS",RedistributionMatchCount=1]
        )))))),
    CareChecks = #table(type table [Check=nullable text,Severity=nullable text,Facility=nullable text,
        MinuteCategory=nullable text,Role=nullable text,Actual=nullable number,Expected=nullable number,Message=nullable text],{}),
    Inputs = [Master=Master,Roles=Roles,Analysis=Analysis,Units=Units,UnitAnalysis=UnitAnalysis,ShiftEnds=ShiftEnds,
        CareDemand=CareRows,CareProfile=CareRows,CareChecks=CareChecks],
    Result = Evaluate(Inputs),
    Demand = Table.Buffer(Result[Demand]),
    Checks = Table.Buffer(Result[Checks]),
    CS = Table.SelectRows(Demand,each [Role]="CS"),
    Care = Table.SelectRows(Demand,each [Role]<>"CS"),
    ExpectedColumns = {"Facility","Role","Week No","FortnightWeek","DayOfWeek","FortnightDayIndex","Shift","DemandHRS","ProductiveHRS","DemandBasis"},
    Cell = (T as table,R as text,D as number,S as text) as record => Table.SelectRows(T,each [Role]=R and [FortnightDayIndex]=D and [Shift]=S){0},
    Sorted = (T as table) as table => Table.Sort(Table.SelectColumns(T,ExpectedColumns),{{"Facility",Order.Ascending},{"Role",Order.Ascending},{"FortnightDayIndex",Order.Ascending},{"Shift",Order.Ascending}}),
    SameCare = (I as record) as logical => Sorted(Table.SelectRows(Evaluate(I)[Demand],each [Role]<>"CS")) = Sorted(Care),
    WithoutCSCell = Table.SelectRows(Master,each not ([Role]="Care Services Employee" and [#"Week No"]=30 and [#"Week Day"]="Monday" and [#"Start Time"]=#time(7,0,0))),
    MissingAllDay = Table.SelectRows(Master,each not ([#"Week No"]=30 and [#"Week Day"]="Monday")),
    SparseWeeks = Table.SelectRows(Master,each not ([Role]="Care Services Employee" and [#"Week No"]=30)),
    ForwardingWarnings = #table(Table.ColumnNames(CareChecks),{
        {"Master Roster Location is outside configured Units","Warning",null,null,null,1,0,"Known excluded-location warning"},
        {"Synthetic global care warning","Warning",null,null,null,1,0,"Keep this global warning"},
        {"Synthetic BD care warning","Warning","BD",null,"RN",1,0,"Keep this enabled-facility warning"},
        {"Synthetic TE care warning","Warning","TE",null,"RN",1,0,"Keep this warning despite the unrelated disabled facility flag"}
    }),
    WarningResult = Evaluate(Change(Inputs,[CareChecks=ForwardingWarnings])),
    ForwardedChecks = Table.Buffer(WarningResult[Checks]),
    Rows = {
        Check("Published common schema is exact",()=>Table.ColumnNames(Demand)=ExpectedColumns),
        Check("Every published value and check is evaluated without cell errors",()=>Table.RowCount(Table.SelectRowsWithErrors(Demand))=0 and Table.RowCount(Table.SelectRowsWithErrors(Checks))=0 and Table.RowCount(ErrorChecks(Checks))=0),
        Check("All three care groups and CS publish unique complete fortnight patterns",()=>Table.RowCount(Demand)=4*42 and Table.RowCount(Table.Distinct(Demand,{"Facility","Role","Week No","DayOfWeek","Shift"}))=4*42 and List.Sort(List.Distinct(Demand[Role]))={"AIN","CS","EN","RN"}),
        Check("CS uses raw net hours and preserves every source cell",()=>Table.RowCount(CS)=42 and List.AllTrue(List.Transform(Table.ToRecords(CS),each Near([DemandHRS],CSAmount([FortnightWeek],[DayOfWeek],List.PositionOf({"AM","PM","NIGHT"},[Shift])+1))))),
        Check("CS fortnight total is the independent raw net-hour sum",()=>Near(List.Sum(CS[DemandHRS]),List.Sum(CSRaw[Shift Net Length]))),
        Check("CS basis is MasterRoster and productive hours remain null",()=>List.Distinct(CS[DemandBasis])={"MasterRoster"} and List.NonNullCount(CS[ProductiveHRS])=0),
        Check("Enabled CS has an explicit passing inclusion diagnostic",()=>let R=Table.SelectRows(Checks,each [Check]="CS effort-analysis inclusion" and [Role]="CS") in Table.RowCount(R)=1 and R{0}[Severity]="Pass" and R{0}[Actual]=1 and R{0}[Expected]=1),
        Check("Care roster demand remains unchanged",()=>List.AllTrue(List.Transform(Table.ToRecords(CareRows),(R)=>Near(Cell(Care,R[Role],R[FortnightDayIndex],if R[Shift]="NS" then "NIGHT" else R[Shift])[DemandHRS],R[RedistributedDemandHours])))),
        Check("Care productive hours apply each source role percentage once",()=>List.AllTrue(List.Transform(Table.ToRecords(CareRows),(R)=>Near(Cell(Care,R[Role],R[FortnightDayIndex],if R[Shift]="NS" then "NIGHT" else R[Shift])[ProductiveHRS],R[RedistributedDemandHours]*R[#"Direct Care %"]))) and List.Distinct(Care[DemandBasis])={"CareMinutes"}),
        Check("The two source weeks remain distinct without averaging",()=>Near(Cell(CS,"CS",1,"AM")[DemandHRS],5.11) and Near(Cell(CS,"CS",8,"AM")[DemandHRS],8.11) and Cell(CS,"CS",1,"AM")[Week No]=30 and Cell(CS,"CS",8,"AM")[Week No]=31),
        Check("Night shift is published as NIGHT",()=>List.Sort(List.Distinct(Demand[Shift]))={"AM","NIGHT","PM"}),
        Check("CSH is excluded even though its own effort flag is true",()=>not List.Contains(Demand[Role],"CSH") and Near(List.Sum(CS[DemandHRS]),List.Sum(CSRaw[Shift Net Length]))),
        Check("False or null CS inclusion warns explicitly and preserves care publication",()=>List.AllTrue(List.Transform({false,null},(Flag)=>let
            A=Table.FromRecords(List.Transform(Table.ToRecords(Analysis),each if [Roles]="CS" then Set(_,"Effort Management Analysis",Flag) else _)),
            I=Change(Inputs,[Analysis=A]), R=Evaluate(I), C=Table.Buffer(R[Checks]),
            Inclusion=Table.SelectRows(C,each [Check]="CS effort-analysis inclusion" and [Role]="CS")
        in Table.RowCount(Table.Buffer(R[Demand]))=126 and SameCare(I)
            and Table.RowCount(Inclusion)=1 and Inclusion{0}[Severity]="Warning"
            and Inclusion{0}[Actual]=0 and Inclusion{0}[Expected]=1))),
        Check("Known out-of-scope care-location warning is not forwarded",()=>Table.RowCount(Table.SelectRows(ForwardedChecks,each [Check]="Master Roster Location is outside configured Units"))=0),
        Check("Other global care warnings remain unchanged",()=>let R=Table.SelectRows(ForwardedChecks,each [Check]="Synthetic global care warning") in Table.RowCount(R)=1 and R{0}=ForwardingWarnings{1}),
        Check("In-scope facility care warnings remain unchanged",()=>let R=Table.SelectRows(ForwardedChecks,each [Check]="Synthetic BD care warning") in Table.RowCount(R)=1 and R{0}=ForwardingWarnings{2}),
        Check("Facility Analysis does not hide other facility care warnings",()=>let R=Table.SelectRows(ForwardedChecks,each [Check]="Synthetic TE care warning") in Table.RowCount(R)=1 and R{0}=ForwardingWarnings{3}),
        Check("Forwarding diagnostics does not change demand or introduce error checks",()=>Sorted(Table.Buffer(WarningResult[Demand]))=Sorted(Demand) and Table.RowCount(Table.SelectRowsWithErrors(ForwardedChecks))=0 and Table.RowCount(ErrorChecks(ForwardedChecks))=0),
        Check("Identical mapping duplicates do not multiply CS demand",()=>let I=Change(Inputs,[Roles=Table.Combine({Roles,Table.SelectRows(Roles,each [Role Group]="CS")})]) in Sorted(Evaluate(I)[Demand])=Sorted(Demand)),
        Check("A second configured CS roster label adds its exact net hours",()=>let Added=Set(Set(CSRaw{0},"Role","Care Services Employee alternate"),"Shift Net Length",1.25), I=Change(Inputs,[Master=Table.Combine({Master,Table.FromRecords({Added},RawColumns)})]) in Near(Cell(Evaluate(I)[Demand],"CS",1,"AM")[DemandHRS],6.36) and SameCare(I)),
        Check("Missing CS source cell becomes a proven zero when facility-day coverage exists",()=>let R=Evaluate(Change(Inputs,[Master=WithoutCSCell]))[Demand] in Table.RowCount(Table.SelectRows(R,each [Role]="CS"))=42 and Cell(R,"CS",1,"AM")[DemandHRS]=0 and Near(Cell(R,"CS",8,"AM")[DemandHRS],8.11)),
        Check("CS sparse first week retains the actual second-week key",()=>let R=Evaluate(Change(Inputs,[Master=SparseWeeks]))[Demand] in List.Sum(Table.SelectRows(R,each [Role]="CS" and [FortnightWeek]=1)[DemandHRS])=0 and Near(Cell(R,"CS",8,"AM")[DemandHRS],8.11)),
        Check("Missing all source coverage for a facility day blocks publication",()=>Reject(Change(Inputs,[Master=MissingAllDay]))),
        Check("Enabled CS without any CS source rows blocks publication",()=>Reject(Change(Inputs,[Master=Table.SelectRows(Master,each [Role]<>"Care Services Employee")]))),
        Check("Negative CS net hours block publication",()=>Reject(Change(Inputs,[Master=Table.Combine({UpdateFirst(CSRaw,"Shift Net Length",-1),CareRaw,HospitalityRaw})]))),
        Check("Null CS net hours block publication",()=>Reject(Change(Inputs,[Master=Table.Combine({UpdateFirst(CSRaw,"Shift Net Length",null),CareRaw,HospitalityRaw})]))),
        Check("Missing care facility field blocks publication",()=>Reject(Change(Inputs,[CareProfile=Table.RemoveColumns(CareRows,{"Facility"})]))),
        Check("Blank care facility key blocks publication",()=>Reject(Change(Inputs,[CareProfile=UpdateFirst(CareRows,"Facility",null)]))),
        Check("Conflicting CS mapping blocks publication",()=>Reject(Change(Inputs,[Roles=Table.Combine({Roles,#table(Table.ColumnNames(Roles),{{"Care Services Employee",null,"CSH","NA",null}})})]))),
        Check("Duplicate analysis keys block publication without an inclusion-control row",()=>let I=Change(Inputs,[Analysis=Table.Combine({Analysis,Table.SelectRows(Analysis,each [Roles]="CS")})]) in Reject(I) and Table.RowCount(Table.SelectRows(Evaluate(I)[Checks],each [Check]="CS effort-analysis inclusion"))=0),
        Check("Missing CS analysis row blocks publication without an inclusion-control row",()=>let I=Change(Inputs,[Analysis=Table.SelectRows(Analysis,each [Roles]<>"CS")]) in Reject(I) and Table.RowCount(Table.SelectRows(Evaluate(I)[Checks],each [Check]="CS effort-analysis inclusion"))=0),
        Check("Ambiguous configured facility mapping blocks publication",()=>Reject(Change(Inputs,[Units=Table.Combine({Units,#table({"Location","Unit"},{{"BD: Beaudesert","TE"}})})]))),
        Check("Unlisted CS locations stay excluded from demand and mapped-row counts without a summary warning",()=>let
            Added=Set(CSRaw{0},"Location","XX: Unlisted"),
            R=Evaluate(Change(Inputs,[Master=Table.Combine({Master,Table.FromRecords({Added},RawColumns)})])),
            C=Table.Buffer(R[Checks]), Stage=Table.SelectRows(C,each [Check]="CS roster-role and source rows")
        in Sorted(R[Demand])=Sorted(Demand) and Table.RowCount(Stage)=1 and Stage{0}[Actual]=42
            and Table.RowCount(Table.SelectRows(C,each [Check]="CS unlisted source locations"))=0),
        Check("Legacy care publication errors block the common publication",()=>Reject(Change(Inputs,[CareChecks=#table(Table.ColumnNames(CareChecks),{{"Synthetic care error","Error","BD",null,"RN",1,0,"Deliberate input error"}})])))
    }
in Table.FromRecords(Rows);
