section CSUnitDemandFixtures;

// Query: CSUnitDemandFixtures
// Purpose: Prepare native acceptance cases for the actual BD generic demand adapter.
// Inputs: Evaluate binds exact production expressions to synthetic import navigation only.
// Output: Pass/fail records for a future explicitly approved SDK run; this file is not a runtime result.
shared CSUnitDemandFixtures = (Evaluate as function) as table => let
    Near=(A as number,B as number) as logical => Number.Abs(A-B)<=0.0000001,
    Change=(Base as record,Fields as record) as record => Record.Combine({Base,Fields}),
    Set=(Row as record,Field as text,Value as any) as record => Record.TransformFields(Row,{{Field,each Value}}),
    UpdateFirst=(T as table,Field as text,Value as any) as table => Table.FromRecords({Set(T{0},Field,Value)} & List.Skip(Table.ToRecords(T))),
    Check=(Name as text,Work as function) as record => let R=try Work(),P=if R[HasError] then false else R[Value]=true
        in [Case=Name,Passed=P,Details=if R[HasError] then R[Error][Message] else if P then null else "Unexpected result"],
    Groups={"RN","AIN","EN","CS"},
    BaseHours=(R as text) as number => if R="RN" then 2 else if R="AIN" then 3 else if R="EN" then 4 else 5,
    Percentage=(R as text) as nullable number => if R="CS" then null else if R="EN" then 0.5 else 1,
    Amount=(R as text,D as number,S as number) as number =>
        if R="CS" and D=3 and S=2 then 0 else BaseHours(R)*(if D<=7 then 1 else 2)+Number.Mod(D-1,7)/10+S/100,
    Handover=Table.FromRecords(List.Combine(List.Transform(Groups,(R)=>List.Combine(List.Transform({1..14},(D)=>
        List.Transform({1..3},(S)=>let H=Amount(R,D,S),W=1+Number.IntegerDivide(D-1,7)
            in [Facility="BD",Role=R,#"Week No"={17,41}{W-1},FortnightWeek=W,DayOfWeek=Number.Mod(D-1,7)+1,
                FortnightDayIndex=D,Shift={"AM","PM","NIGHT"}{S-1},DemandHRS=H,
                ProductiveHRS=if R="CS" then null else H*Percentage(R),DemandBasis=if R="CS" then "MasterRoster" else "CareMinutes"]
    )))))),
    HandoverChecks=#table({"Check","Severity","Facility","MinuteCategory","Role","Actual","Expected","Message"},
        {{"Synthetic generic handover validated","Pass","BD",null,null,168,168,"Synthetic pass"}}),
    Navigation=(D as table,C as table) as table => #table({"Item","Kind","Data"},{{"RoleDemand_TABLE","Table",D},{"RoleDemand_CHECK","Table",C}}),
    Calendar=(Days as number) as table => Table.FromRecords(List.Combine(List.Transform(List.Numbers(0,Days),(D)=>
        List.Combine(List.Transform({1..3},(S)=>List.Transform(Groups,(R)=>
            [Date=Date.AddDays(#date(2026,9,28),D),Day=D+1,Shifts={"AM","PM","NIGHT"}{S-1},Period=D*3+S,RolesList=R])))))),
    Timing=Table.FromRecords(List.Combine(List.Transform(Groups,(R)=>List.Transform({1..3},(S)=>
        [Role=R,ShiftPeriod={"AM","PM","NIGHT"}{S-1},StartTime={0.25,14/24,21/24}{S-1},
            StartDay={0.25,14/24,21/24}{S-1},EndDay={14.25/24,21.25/24,0.25}{S-1},DurationOfShifts={8.25,7.25,9}{S-1}])))),
    Settings=(Days as number,Cal as nullable table,Times as nullable table) as table => #table({"Item","Kind","Data"},{
        {"ShiftDuration","Table",#table({"ShiftDuration"},{{7.6}})},
        {"ShiftStart","Table",#table({"StartAM","StartPM","StartNIGHT"},{{0.25,14/24,21/24}})},
        {"PermutationDimensions","Table",if Cal=null then Calendar(Days) else Cal},
        {"ShiftPeriod","Table",if Times=null then Timing else Times}}),
    Roles=#table({"Roster Roles","DC Role","Role Group","DC Category","Direct Care %"},{
        {"Registered Nurse","RN","RN","RN",1},{"Assistant in Nursing","AIN","AIN","OTHER",1},
        {"Enrolled Nurse","EN","EN","OTHER",0.5},{"Care Services Employee",null,"CS","NA",null}}),
    Analysis=#table({"Roles","Effort Management Analysis","Leave Balance Analysis"},List.Transform(Groups,each {_,true,true})),
    Facilities=#table({"Title","EffortManagementAnalysis","LeaveBalanceAnalysis"},{{"BD",true,false},{"TE",true,false}}),
    Inputs=[DemandNavigation=Navigation(Handover,HandoverChecks),Settings=Settings(28,null,null),Roles=Roles,Analysis=Analysis,Facilities=Facilities,
        PathNavigation=#table({"Item","Kind","Data"},{{"CentriSyncPaths","Table",#table({"SharepointRootUrl","SyncedFolderRootPath"},{{"https://fixture.invalid/ResidentialCare","Z:\ResidentialCare"}})}}),
        PathInput=#table({"FilePath"},{{"Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD\1. Input\[2-DemandExtract.xlsx]Path"}})],
    WithHandover=(D as table,C as table) as record => Change(Inputs,[DemandNavigation=Navigation(D,C)]),
    WithSettings=(Cal as nullable table,Times as nullable table) as record => Change(Inputs,[Settings=Settings(28,Cal,Times)]),
    Result=Evaluate(Inputs),Demand=Table.Buffer(Result[Demand]),Checks=Table.Buffer(Result[Checks]),
    Pattern=Table.Buffer(Result[Pattern]),Cells=Table.Buffer(Result[Cells]),Reconciliation=Table.Buffer(Result[Reconciliation]),
    ExpectedColumns={"Date","Day","Shift","Period","Role","StartTime","EndTime","Unit","Facility","DemandFTE","DemandHRS","DurationOfShifts"},
    Sorted=(T as table) as table => Table.Sort(Table.SelectColumns(T,ExpectedColumns),{{"Period",Order.Ascending},{"Role",Order.Ascending}}),
    Reject=(I as record) as logical => let R=Evaluate(I),C=try Table.Buffer(R[Checks]),D=try Table.RowCount(Table.Buffer(R[Demand]))
        in not C[HasError] and Table.RowCount(Table.SelectRows(C[Value],each [Severity]="Error"))>0 and D[HasError],
    ValidHours=(T as table) as logical => List.AllTrue(List.Transform(Table.ToRecords(T),each
        Near([DemandHRS],Amount([Role],Number.Mod([Day]-1,14)+1,List.PositionOf({"AM","PM","NIGHT"},[Shift])+1)))),
    ExpectedTotal=(R as text,Days as number) as number => List.Sum(List.Combine(List.Transform(List.Numbers(0,Days),(D)=>List.Transform({1..3},(S)=>Amount(R,Number.Mod(D,14)+1,S))))),
    AddCheck=(Facility as nullable text,Severity as text) as table => Table.Combine({HandoverChecks,
        #table(Table.ColumnNames(HandoverChecks),{{"Synthetic handover issue",Severity,Facility,null,"CS",1,0,"Deliberate issue"}})}),
    CSRows=Table.SelectRows(Handover,each [Role]="CS"),CareRows=Table.SelectRows(Handover,each [Role]<>"CS"),
    Rows={
        Check("Twelve-column downstream publication is unchanged",()=>Table.ColumnNames(Demand)=ExpectedColumns),
        Check("All four roles have one row per Settings period",()=>Table.RowCount(Demand)=336 and Table.RowCount(Table.Distinct(Demand,{"Role","Period"}))=336 and List.Sort(List.Distinct(Demand[Role]))={"AIN","CS","EN","RN"}),
        Check("Publication checks reconciliation and legacy helper aliases are forced without errors",()=>List.AllTrue(List.Transform(
            {Demand,Checks,Pattern,Cells,Reconciliation,Result[Profile],Result[Exclusions],Result[CalendarCheck]},
            each Table.RowCount(Table.SelectRowsWithErrors(Table.Buffer(_)))=0))
            and Table.RowCount(Table.SelectRows(Checks,each [Severity]="Error"))=0),
        Check("All care and CS hour cells retain the exact unequal source pattern",()=>ValidHours(Demand)),
        Check("Source week numbers are preserved rather than renumbered",()=>List.Sort(List.Distinct(Cells[Week No]))={17,41} and Table.RowCount(Table.SelectRows(Cells,each ([FortnightWeek]=1 and [Week No]<>17) or ([FortnightWeek]=2 and [Week No]<>41)))=0),
        Check("CS productive hours remain null in source and grouped pattern",()=>List.NonNullCount(Table.SelectRows(Cells,each [Role]="CS")[ProductiveHRS])=0 and List.NonNullCount(Table.SelectRows(Pattern,each [Role]="CS")[ProductiveHRS])=0),
        Check("Care productive hours remain numeric and unchanged",()=>List.AllTrue(List.Transform(Table.ToRecords(Table.SelectRows(Cells,each [Role]<>"CS")),each Value.Is([ProductiveHRS],type number) and Near([ProductiveHRS],[DemandHRS]*Percentage([Role]))))),
        Check("CS roster hours are conserved across two planning fortnights",()=>Near(List.Sum(Table.SelectRows(Demand,each [Role]="CS")[DemandHRS]),ExpectedTotal("CS",28))),
        Check("FTE uses each actual shift duration and round-trips to hours",()=>List.AllTrue(List.Transform(Table.ToRecords(Demand),each Near([DemandFTE],[DemandHRS]/[DurationOfShifts]) and Near([DemandFTE]*[DurationOfShifts],[DemandHRS])))),
        Check("Overnight end timestamps advance exactly one date",()=>List.AllTrue(List.Transform(Table.ToRecords(Demand),each Near(Duration.TotalHours([EndTime]-[StartTime]),[DurationOfShifts]) and (if [Shift]="NIGHT" then Date.From([EndTime])=Date.AddDays([Date],1) else Date.From([EndTime])=[Date])))),
        Check("Every full-fortnight reconciliation passes",()=>Table.RowCount(Reconciliation)=8 and List.Distinct(Reconciliation[Status])={"Pass"}),
        Check("One through thirty-five day horizons retain exact source day totals",()=>List.AllTrue(List.Transform({1,7,14,15,21,28,35},(Days)=>let R=Evaluate(Change(Inputs,[Settings=Settings(Days,null,null)])),D=Table.Buffer(R[Demand]),C=Table.Buffer(R[Reconciliation]) in Table.RowCount(D)=Days*3*4 and ValidHours(D) and List.Distinct(C[Status])={"Pass"} and List.AllTrue(List.Transform(Groups,(Role)=>Near(List.Sum(Table.SelectRows(D,each [Role]=Role)[DemandHRS]),ExpectedTotal(Role,Days))))))),
        Check("Final partial fortnight uses actual source days without proration",()=>let R=Evaluate(Change(Inputs,[Settings=Settings(15,null,null)])),C=Table.SelectRows(R[Reconciliation],each [PlanningFortnight]=2) in Table.RowCount(C)=4 and List.AllTrue(List.Transform(Table.ToRecords(C),each [ExpectedCellCount]=3 and Near([ExpectedRosterHours],ExpectedTotal([Role],1))))),
        Check("Folder-derived full BD unit key is retained",()=>Result[Facility]="BD" and Result[UnitPath]="Z:\ResidentialCare\CLIENT\DATExx-Whiddon\UNITS\BD" and List.Distinct(Demand[Facility])={"BD"} and List.Distinct(Demand[Unit])={"BD"}),
        Check("Other facility and BD-prefix alternatives are excluded",()=>let Other=Table.Combine({Handover,Table.TransformColumns(Handover,{{"Facility",each "TE"}}),Table.TransformColumns(Handover,{{"Facility",each "BD-OTHER"}})}) in Sorted(Evaluate(WithHandover(Other,HandoverChecks))[Demand])=Sorted(Demand)),
        Check("Global upstream handover errors block publication",()=>Reject(WithHandover(Handover,AddCheck(null,"Error")))),
        Check("BD upstream handover errors block publication",()=>Reject(WithHandover(Handover,AddCheck("BD","Error")))),
        Check("Other-facility upstream issues cannot block BD",()=>Sorted(Evaluate(WithHandover(Handover,AddCheck("TE","Error")))[Demand])=Sorted(Demand)),
        Check("An empty handover check publication is not a runtime pass",()=>Reject(WithHandover(Handover,Table.FirstN(HandoverChecks,0)))),
        Check("Non-null CS productive hours fail visibly",()=>Reject(WithHandover(Table.Combine({UpdateFirst(CSRows,"ProductiveHRS",1),CareRows}),HandoverChecks))),
        Check("Null care productive hours fail visibly",()=>Reject(WithHandover(Table.Combine({CSRows,UpdateFirst(CareRows,"ProductiveHRS",null)}),HandoverChecks))),
        Check("An unsupported demand basis fails visibly",()=>Reject(WithHandover(UpdateFirst(Handover,"DemandBasis","Unknown"),HandoverChecks))),
        Check("Duplicate handover cells block publication",()=>Reject(WithHandover(Table.Combine({Handover,Table.FirstN(Handover,1)}),HandoverChecks))),
        Check("Missing CS source pattern cells cannot become invented zeroes",()=>Reject(WithHandover(Table.Combine({Table.Skip(CSRows,1),CareRows}),HandoverChecks))),
        Check("Stale Settings without CS fails role coverage",()=>Reject(WithSettings(Table.SelectRows(Calendar(28),each [RolesList]<>"CS"),null))),
        Check("One missing CS Settings period remains a coverage failure",()=>Reject(WithSettings(Table.SelectRows(Calendar(28),each not ([RolesList]="CS" and [Period]=2)),null))),
        Check("Missing CS shift definitions block publication",()=>Reject(WithSettings(null,Table.SelectRows(Timing,each not ([Role]="CS" and [ShiftPeriod]="NIGHT"))))),
        Check("Mismatched duration and timestamp spans block publication",()=>Reject(WithSettings(null,UpdateFirst(Timing,"DurationOfShifts",999)))),
        Check("Missing authoritative DemandHRS cannot fall back to FTE",()=>Reject(WithHandover(Table.AddColumn(Table.RemoveColumns(Handover,{"DemandHRS"}),"FTE",each 999),HandoverChecks))),
        Check("Invalid handover hours fail without defaults",()=>List.AllTrue(List.Transform({null,"7.6",-1,#nan,#infinity},(H)=>Reject(WithHandover(UpdateFirst(Handover,"DemandHRS",H),HandoverChecks))))),
        Check("Additional display FTE cannot change authoritative hours",()=>Sorted(Evaluate(WithHandover(Table.AddColumn(Handover,"FTE",each 999),HandoverChecks))[Demand])=Sorted(Demand)),
        Check("Legacy-only navigation is not a generic handover substitute",()=>Reject(Change(Inputs,[DemandNavigation=#table({"Item","Kind","Data"},{{"MinuteWorkersFTE_TABLE","Table",Handover},{"MinuteWorkersFTE_CHECK","Table",HandoverChecks}})]))),
        Check("Source row order cannot average or renumber source weeks",()=>Sorted(Evaluate(WithHandover(Table.ReverseRows(Handover),HandoverChecks))[Demand])=Sorted(Demand)),
        Check("Upstream warnings remain visible without altering demand",()=>let R=Evaluate(WithHandover(Handover,AddCheck("BD","Warning"))) in Sorted(R[Demand])=Sorted(Demand) and Table.RowCount(Table.SelectRows(R[Checks],each [Severity]="Warning" and [Check]="Synthetic handover issue"))=1)
    }
in Table.FromRecords(Rows);
