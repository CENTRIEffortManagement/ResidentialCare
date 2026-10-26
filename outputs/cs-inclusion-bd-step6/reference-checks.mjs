// Independent business-rule examples only. This script does not execute M or read workbooks.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const scratch=path.dirname(fileURLToPath(import.meta.url));
const groups=['RN','AIN','EN','CS'];
const text=value=>value===null||value===undefined||String(value).trim()===''?null:String(value).trim();
const key=(id,facility)=>JSON.stringify([text(id),facility]);
let assertions=0;
const cases=[];
const equal=(actual,expected,label)=>{assertions++;assert.deepEqual(actual,expected,label);};
const throws=(fn,pattern,label)=>{assertions++;assert.throws(fn,pattern,label);};
function scenario(name,fn){fn();cases.push(name);}

// Inputs here have already resolved canonical facility identity. No facility alias policy is invented.
// The raw-published key set also contains deliberately excluded agency workers.
function bridge(allocations,rawPublished,unit='BD') {
    const publishedKeys=new Set(rawPublished.map(row=>key(row.ID,row.Location===undefined?row.Facility:text(row.Location)?.slice(0,2).toUpperCase())));
    const publishedCanonicalKeys=new Set();
    for(const row of allocations) {
        const id=text(row.ID),rawID=text(row.RawID??row.ID),sourceFacility=row.SourceFacility??row.Facility;
        if(id!==null&&(publishedKeys.has(key(id,sourceFacility))||publishedKeys.has(key(rawID,sourceFacility))))publishedCanonicalKeys.add(key(id,sourceFacility));
    }
    const workers=new Map();
    for(const row of allocations) {
        const id=text(row.ID),rawID=text(row.RawID??row.ID),name=text(row.Name);
        if(text(row.Role)!=='CS'||row.Facility!==unit||id===null)continue;
        const workerKey=key(id,row.Facility);
        const sourceFacility=row.SourceFacility??row.Facility;
        if(publishedCanonicalKeys.has(key(id,sourceFacility)))continue;
        if(row.RoleShiftGroup!=='CS')throw new Error('Saved RoleShift does not resolve to CS');
        if(name===null)throw new Error('Missing assigned CS name');
        const baseName=row.RoleCount>1&&name.endsWith(' (CS)')?name.slice(0,-5):name;
        const previous=workers.get(workerKey);
        if(previous&&previous.Name!==baseName)throw new Error('Ambiguous assigned CS name');
        workers.set(workerKey,{ID:id,Facility:row.Facility,Name:baseName,Role:'CS',InAllocation:true});
    }
    return [...workers.values()];
}

function reconcile(publishedPreferred,bridgeWorkers,records,hr,rosterStart='2026-10-26') {
    const candidates=new Map(publishedPreferred.map(row=>[key(row.ID,row.Facility),{...row,InAllocation:true}]));
    for(const row of bridgeWorkers)if(!candidates.has(key(row.ID,row.Facility)))candidates.set(key(row.ID,row.Facility),{...row});
    for(const row of records) {
        const workerKey=key(row.ID,row.Facility),previous=candidates.get(workerKey);
        candidates.set(workerKey,previous?{...previous,InRecords:true}:{...row,InAllocation:false,InRecords:true});
    }
    const result=[];
    for(const row of candidates.values()) {
        const employment=hr.find(item=>text(item.ID)===text(row.ID));
        if(employment?.Termination&&employment.Termination<rosterStart)continue;
        if(!row.InAllocation&&!employment)throw new Error('Availability worker has no current worker match');
        if(employment?.EligibleRowCount>1)throw new Error('Current worker record ambiguous');
        const role=row.Role??employment?.Role;
        if(!groups.includes(role))continue;
        result.push({...row,Role:role,PreferredRole:role,ContractHours:employment?.Hours??null});
    }
    return result;
}

function shares(allocations,unit='BD') {
    const grouped=new Map();
    for(const row of allocations) {
        if(row.Facility!==unit||!groups.includes(row.Role))continue;
        if(!Number.isFinite(row.Hours)||row.Hours<0)throw new Error('Invalid allocation hours');
        const groupKey=JSON.stringify([text(row.ID),row.Facility,row.Role]);
        const group=grouped.get(groupKey)??{ID:text(row.ID),Facility:row.Facility,Role:row.Role,Hours:0};
        group.Hours+=row.Hours;grouped.set(groupKey,group);
    }
    return [...grouped.values()].map(row=>{
        const total=[...grouped.values()].filter(other=>key(other.ID,other.Facility)===key(row.ID,row.Facility)).reduce((sum,other)=>sum+other.Hours,0);
        if(total<=0)throw new Error('Approved-role allocation total must be positive');
        return {...row,Share:row.Hours/total};
    });
}

const allocation=(ID,Role='CS',Facility='BD',Name=`Worker ${ID}`,Hours=7.5)=>({ID,Role,Facility,Name,Hours,RoleShiftGroup:Role,RoleCount:1});
scenario('filled CS only; CSH, care, other facility and unassigned rows are excluded',()=>{
    const rows=[allocation('100'),allocation('101','CSH'),allocation('102','AIN'),allocation('103','CS','TE'),allocation('  '),allocation(null)];
    equal(bridge(rows,[]),[{ID:'100',Facility:'BD',Name:'Worker 100',Role:'CS',InAllocation:true}],'Exactly one eligible CS employee');
});
scenario('repeated allocations create one employee/facility candidate',()=>{
    const rows=Array.from({length:84},()=>allocation(100));
    equal(bridge(rows,[]).length,1,'Narrow deduplication is independent of shift count');
    equal(rows.length,84,'Allocation evidence remains intact');
});
scenario('a zero-hour allocation fact does not erase an otherwise valid worker',()=>{
    const rows=[allocation('100','CS','BD','Worker 100',0),allocation('100')];
    equal(bridge(rows,[]).length,1,'Identity selection does not discard the employee because one fact is zero');
    equal(shares(rows)[0].Share,1,'The existing positive-total check still passes for this valid employee');
});
scenario('published AIN plus CS allocation retains the existing preferred resource',()=>{
    const raw=[{ID:'100',Facility:'BD',Name:'Worker 100',Role:'AIN'}];
    const additions=bridge([allocation('100','CS','BD','Worker 100 (CS)')],raw);
    equal(additions,[],'No second CS candidate for the existing raw-roster employee');
    const workers=reconcile(raw,additions,[],[{ID:'100',Hours:76,Role:'AIN'}]);
    equal(workers.length,1,'One preferred resource');
    equal([workers[0].Name,workers[0].PreferredRole,workers[0].ContractHours],['Worker 100','AIN',76],'Name, role and employee contract remain unchanged');
});
scenario('published agency exclusion is retained',()=>{
    const raw=[{ID:'100',Facility:'BD',Name:'Agency worker',Role:'AIN',EmploymentType:'Agency'}];
    equal(bridge([allocation('100')],raw),[],'CS allocation does not resurrect a deliberately excluded published agency worker');
});
scenario('saved canonical and raw employee aliases both suppress existing employees',()=>{
    const row={...allocation('100'),RawID:'HR100'};
    equal(bridge([row],[{ID:'HR100',Facility:'BD'}]),[],'Raw employee alias prevents a second resource');
    equal(bridge([row],[{ID:'100',Facility:'BD'}]),[],'Canonical code also prevents a second resource');
});
scenario('published agency identified through raw alias remains excluded',()=>{
    equal(bridge([{...allocation('100'),RawID:'HR100'}],[{ID:'HR100',Facility:'BD',EmploymentType:'Agency'}]),[],'Agency cannot be resurrected through a different canonical code');
});
scenario('raw published location variants cannot resurrect excluded employees',()=>{
    equal(bridge([allocation('100')],[{ID:'100',Location:'  bd - old source spelling  '}]),[],'Presence guard uses raw source-prefix identity before legacy location exclusion');
});
scenario('agency plus location variant is blocked through a cross-role alias',()=>{
    const rows=[{...allocation('100','AIN'),RawID:'900'},{...allocation('100'),RawID:'100'}];
    equal(bridge(rows,[{ID:'900',Location:'BD - excluded variant',EmploymentType:'Agency'}]),[],'Both original exclusions remain effective through canonical identity');
});
scenario('an existing published alias from another saved role blocks a duplicate CS resource',()=>{
    const rows=[{...allocation('100','AIN'),RawID:'900'},{...allocation('100'),RawID:'100'}];
    equal(bridge(rows,[{ID:'900',Facility:'BD',Role:'AIN'}]),[],'All saved aliases for the canonical employee participate, including other roles');
});
scenario('one matched CS alias suppresses every candidate for that canonical employee',()=>{
    const rows=[{...allocation('100'),RawID:'900'},{...allocation('100'),RawID:'100'}];
    equal(bridge(rows,[{ID:'900',Facility:'BD'}]),[],'Per-row aliases cannot split one canonical employee into a new resource');
});
scenario('anti-join retains existing source-facility grain',()=>{
    equal(bridge([{...allocation('100'),SourceFacility:'XX'}],[{ID:'100',Facility:'XX'}]),[],'Canonical BD scope does not erase raw source-facility identity');
});
scenario('genuine saved RoleShift mapping is required',()=>{
    throws(()=>bridge([{...allocation('100'),RoleShiftGroup:'CSH'}],[]),/RoleShift/,'A CS table label cannot override a CSH raw role');
    throws(()=>bridge([{...allocation('100'),RoleShiftGroup:null}],[]),/RoleShift/,'Missing raw role mapping is not invented');
});
scenario('only a verified generated CS suffix is removed',()=>{
    equal(bridge([{...allocation('100','CS','BD','Worker 100 (CS)'),RoleCount:2}],[])[0].Name,'Worker 100','Multi-role generated suffix is removed once');
    equal(bridge([allocation('100','CS','BD','Worker 100 (CS)')],[])[0].Name,'Worker 100 (CS)','Single-role legitimate name is unchanged');
    equal(bridge([{...allocation('100','CS','BD','Worker 100 (cs)'),RoleCount:2}],[])[0].Name,'Worker 100 (cs)','No invented case-insensitive name rewrite');
});
scenario('same ID at another facility does not block BD membership',()=>equal(bridge([allocation('100')],[{ID:'100',Facility:'TE'}]).length,1,'Anti-join key includes facility'));
scenario('text identifiers retain leading zeros',()=>equal(bridge([allocation('00100')],[{ID:'100',Facility:'BD'}])[0].ID,'00100','No numeric ID coercion'));
scenario('assigned employee missing a name fails',()=>throws(()=>bridge([allocation('100','CS','BD',' ')],[]),/Missing assigned/,'Missing identity is not fabricated'));
scenario('conflicting names for a new CS ID fail',()=>throws(()=>bridge([allocation('100'),allocation('100','CS','BD','Different name')],[]),/Ambiguous/,'Conflicting names do not multiply workers'));
scenario('availability-only CS membership is upgraded to both without duplication',()=>{
    const additions=bridge([allocation('100')],[]);
    const workers=reconcile([],additions,[{ID:'100',Facility:'BD',Name:'Records name',Role:'CS'}],[{ID:'100',Hours:38}]);
    equal(workers.length,1,'One employee');
    equal([workers[0].InAllocation,workers[0].InRecords,workers[0].Name,workers[0].ContractHours],[true,true,'Worker 100',38],'Allocation identity and existing HR contract survive');
});
scenario('strict-before-start termination exclusion and boundary are unchanged',()=>{
    const additions=bridge([allocation('100'),allocation('101'),allocation('102')],[]);
    const hr=[{ID:'100',Termination:'2026-10-25'},{ID:'101',Termination:'2026-10-26'},{ID:'102',Termination:'2026-11-01'}];
    equal(reconcile([],additions,[],hr).map(row=>row.ID),['101','102'],'On-start and future termination are retained');
});
scenario('allocation evidence keeps existing missing-HR fallback',()=>{
    const workers=reconcile([],bridge([allocation('100')],[]),[],[]);
    equal([workers.length,workers[0].ContractHours],[1,null],'No new mandatory-HR rule or invented contract');
});
scenario('records-only missing HR remains blocking',()=>throws(()=>reconcile([],[],[{ID:'100',Facility:'BD',Name:'Worker 100',Role:'CS'}],[]),/no current worker match/,'Existing membership rule'));
scenario('ambiguous HR remains blocking',()=>throws(()=>reconcile([],bridge([allocation('100')],[]),[],[{ID:'100',EligibleRowCount:2}]),/ambiguous/,'Existing ambiguity rule'));
scenario('wholly unavailable CS remains an identifiable employee',()=>{
    const workers=reconcile([],bridge([allocation('100')],[]),[{ID:'100',Facility:'BD',Name:'Worker 100',Role:'CS',UnavailableAllRoster:true}],[{ID:'100',Hours:76}]);
    equal(workers.length,1,'Availability hours do not determine staff-list membership');
});
scenario('HR presence alone does not create membership',()=>equal(reconcile([],[],[],[{ID:'100',Role:'CS',Hours:76}]),[],'HR is enrichment rather than an extra population source'));
function appliedCasualCap(contractHours,sourceCap,maximum,issue=false) {
    if(contractHours!==null&&contractHours!==0)return null;
    return issue||sourceCap===null?maximum:Math.min(sourceCap,maximum);
}
scenario('positive HR contract preserves precedence even for a casual employee',()=>equal(appliedCasualCap(76,2,20),null,'No casual roster-cap replacement for a positive contract'));
scenario('missing or zero contract keeps established roster-cap fallback',()=>{
    equal(appliedCasualCap(null,12,20),12,'Resolved employee roster cap');
    equal(appliedCasualCap(0,30,20),20,'Settings maximum still limits the cap');
    equal(appliedCasualCap(null,null,20),20,'Missing lookup uses the same maximum');
    equal(appliedCasualCap(0,12,20,true),20,'Invalid lookup retains fallback rather than fabricating contract hours');
});
scenario('AIN and CS shares reconcile without duplicate employee contracts',()=>{
    const values=shares([allocation('100','AIN','BD','Worker 100 (AIN)',60),allocation('100','CS','BD','Worker 100 (CS)',40),allocation('100','CSH','BD','Worker 100 (CSH)',999)]);
    equal(values.map(row=>[row.Role,row.Share]),[['AIN',0.6],['CS',0.4]],'CS participates and CSH remains excluded');
    equal(values.reduce((sum,row)=>sum+row.Share,0),1,'Role shares sum to one');
    const published=[{ID:'100',Facility:'BD',Name:'Worker 100',Role:'AIN'}];
    equal(reconcile(published,bridge([allocation('100')],published),[],[{ID:'100',Hours:76}]).map(row=>row.ContractHours),[76],'One unchanged employee contract, rather than one per allocation role');
});
scenario('three roles and duplicate allocation facts still reconcile',()=>{
    const values=shares([allocation('100','AIN','BD',undefined,10),allocation('100','AIN','BD',undefined,10),allocation('100','RN','BD',undefined,10),allocation('100','CS','BD',undefined,10)]);
    equal(values.map(row=>[row.Role,row.Share]),[['AIN',0.5],['RN',0.25],['CS',0.25]],'Group role hours before division');
});
scenario('role shares are separated by facility',()=>{
    equal(shares([allocation('100','AIN','BD',undefined,10),allocation('100','CS','TE',undefined,90)])[0].Share,1,'TE hours cannot dilute BD');
});
scenario('invalid role-share hours continue to fail',()=>{
    throws(()=>shares([allocation('100','CS','BD',undefined,-1)]),/Invalid/,'Negative hours');
    throws(()=>shares([allocation('100','CS','BD',undefined,0)]),/positive/,'Zero total');
});

const day=86400000;
function calendar(start,days,roleList,definitions) {
    const rows=[];
    for(const role of roleList)for(let offset=0;offset<days;offset++)for(const definition of definitions.filter(item=>item.Role===role)) {
        const date=Date.parse(`${start}T00:00:00Z`)+offset*day;
        const endHour=definition.End>definition.Start?definition.End:definition.End+24;
        rows.push({Role:role,Date:new Date(date).toISOString().slice(0,10),Shift:definition.Shift,Start:new Date(date+definition.Start/24*day).toISOString(),End:new Date(date+endHour/24*day).toISOString()});
    }
    return rows;
}
const shiftDefinitions=groups.flatMap(Role=>[{Role,Shift:'AM',Start:7,End:15},{Role,Shift:'PM',Start:15,End:23},{Role,Shift:'NIGHT',Start:23,End:7}]);
scenario('CS receives every dated Settings shift with overnight timestamp coverage',()=>{
    const rows=calendar('2026-10-26',28,groups,shiftDefinitions);
    const cs=rows.filter(row=>row.Role==='CS');
    equal([rows.length,cs.length],[336,84],'Four groups, 28 dates and three shifts');
    equal(new Set(cs.map(row=>`${row.Date}|${row.Shift}`)).size,84,'Unique CS calendar keys');
    equal(cs[2],{Role:'CS',Date:'2026-10-26',Shift:'NIGHT',Start:'2026-10-26T23:00:00.000Z',End:'2026-10-27T07:00:00.000Z'},'Night crosses midnight');
    equal(cs.at(-1).End,'2026-11-23T07:00:00.000Z','The final night retains its full endpoint');
    equal(calendar('2026-10-26',28,['CSH'],shiftDefinitions),[],'No CSH calendar lane');
});
scenario('missing CS Settings coverage is visible before publication',()=>{
    const rows=calendar('2026-10-26',28,groups,shiftDefinitions.filter(row=>row.Role!=='CS'));
    const eligibleWorkers=[{ID:'100',Role:'CS'},{ID:'101',Role:'AIN'}];
    equal(eligibleWorkers.filter(worker=>!rows.some(row=>row.Role===worker.Role)).map(row=>row.ID),['100'],'Required role-coverage gate has a concrete missing CS worker');
});

const summary={status:'Passed independent reference examples',referenceCases:cases.length,referenceAssertions:assertions,cases,nativeEvaluationPerformed:false,productionMExecuted:false,timingMeasured:false,limitation:'Reference JavaScript examples do not execute production M or verify saved workbook data/bindings.'};
fs.writeFileSync(path.join(scratch,'reference-summary.json'),JSON.stringify(summary,null,2));
console.log(JSON.stringify(summary,null,2));
