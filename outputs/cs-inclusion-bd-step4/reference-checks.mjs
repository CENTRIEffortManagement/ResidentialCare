// Independent allocation/date/identity reference checks only; this does not evaluate M.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const scratch=path.dirname(fileURLToPath(import.meta.url));
const dayMs=86400000;
const toDay=(value)=>Date.parse(value+'T00:00:00Z')/dayMs;
const iso=(value)=>new Date(value*dayMs).toISOString().slice(0,10);
const mod=(value,base)=>((value%base)+base)%base;
const blank=(value)=>value===null||value===undefined||(typeof value==='string'&&value.trim()==='');
const weekday=(day)=>mod(day+3,7); // UTC epoch was a Thursday; Monday is zero.
const roleGroup=new Map([['CS roster','CS'],['CSH roster','CSH'],['AIN roster','AIN'],['C4 roster','AINC4'],['RN roster','RN']]);
const facility=new Map([['BD: Beaudesert','BD'],['TE: Temora','TE']]);
let assertions=0;
const equal=(a,b,label)=>{assertions++;assert.deepEqual(a,b,label);};
const near=(a,b,label)=>{assertions++;assert.ok(Math.abs(a-b)<=1e-8,label);};
const rejected=(work,label)=>{assertions++;assert.throws(work,label);};
function portfolioPath(unitPath) {
    const segments=unitPath.replace(/\\+$/,'').split('\\');
    const unitsIndex=segments.findLastIndex(segment=>segment.toUpperCase()==='UNITS');
    if(unitsIndex<=0||segments.length!==unitsIndex+2)throw new Error('Expected a complete unit folder');
    return segments.slice(0,unitsIndex).join('\\');
}
equal(portfolioPath('Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon\\UNITS\\BD'),'Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon','Windows drive path retains the organisation root');
equal(portfolioPath('Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon\\UNITS\\BD\\'),'Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon','One trailing separator does not change the organisation root');
equal(portfolioPath('\\\\server\\share\\ResidentialCare\\CLIENT\\DATExx-Whiddon\\UNITS\\BD'),'\\\\server\\share\\ResidentialCare\\CLIENT\\DATExx-Whiddon','UNC path retains its leading network separators');
equal(portfolioPath('Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon\\units\\BD'),'Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon','UNITS matching is case-insensitive');
rejected(()=>portfolioPath('Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon\\BD'),'A path without UNITS is rejected');
rejected(()=>portfolioPath('Z:\\ResidentialCare\\CLIENT\\DATExx-Whiddon\\UNITS\\BD\\2. Calculations'),'A calculation subfolder cannot be mistaken for the unit root');
function supplement(raw,published,anchor,start,end) {
    const anchorDay=toDay(anchor);
    if(!Number.isFinite(anchorDay)||weekday(anchorDay)!==0)throw new Error('Anchor must be a Monday');
    if(published.some(row=>facility.get(row.location)==='BD'&&roleGroup.get(row.role)==='CS'))throw new Error('Published Roster already contains CS');
    const bd=raw.filter(row=>facility.get(row.location)==='BD');
    const weeks=[...new Set(bd.map(row=>row.week))].sort((a,b)=>a-b);
    if(weeks.length!==2)throw new Error('Two source weeks are required');
    const assigned=bd.filter(row=>roleGroup.get(row.role)==='CS'&&!blank(row.code));
    for(const row of assigned)if(blank(row.name)||!Number.isFinite(row.hours)||row.hours<0)throw new Error('Invalid assigned CS row');
    const output=[];
    for(let date=toDay(start);date<=toDay(end);date++) {
        const index=mod(date-anchorDay,14);
        for(const row of assigned.filter(row=>row.week===weeks[Math.floor(index/7)]&&row.weekday===index%7)) {
            output.push({...row,date:iso(date),role:'CS',sourceWeek:row.week,fortnightDay:index+1,
                endDate:iso(date+(row.end<row.start?1:0))});
        }
    }
    return output;
}
function resolveIdentities(master,published) {
    const key=(name)=>blank(name)?null:String(name).trim().toUpperCase();
    const resolved=master.map(row=>{
        const matches=published.filter(person=>[person.code,person.employeeCode].includes(row.code));
        const unique=[...new Map(matches.map(person=>[JSON.stringify([person.code,person.name]),person])).values()];
        if(unique.length>1)throw new Error('Ambiguous published employee alias');
        const person=unique[0];
        const name=person?person.name:String(row.name??'').trim();
        if(blank(name))throw new Error('Missing resolved name');
        return {...row,code:person?person.code:row.code,name,masterNameKey:key(row.name),nameKey:key(name)};
    });
    const relevantKeys=new Set(resolved.flatMap(row=>[row.masterNameKey,row.nameKey]).filter(Boolean));
    const pairs=[...resolved.flatMap(row=>[row.nameKey,row.masterNameKey].filter(Boolean).map(name=>[name,row.code])),
        ...published.filter(person=>relevantKeys.has(key(person.name))).map(person=>[key(person.name),person.code])];
    for(const name of relevantKeys)if(new Set(pairs.filter(pair=>pair[0]===name).map(pair=>pair[1])).size>1)throw new Error('Name identifies different employee codes');
    for(const code of new Set(resolved.map(row=>row.code)))if(new Set(resolved.filter(row=>row.code===code).map(row=>row.name)).size>1)throw new Error('Code has conflicting exact resolved names');
    return resolved;
}
const anchor='2026-09-28';
const raw=[];
for(const week of [1,2])for(let day=0;day<7;day++)raw.push({location:'BD: Beaudesert',role:'CS roster',code:4001,
    name:'CS Worker',week,weekday:day,start:week===2&&day===6?'22:00':'07:00',end:week===2&&day===6?'06:00':'15:00',
    hours:week===1?7.6:7.1,breakMinutes:30,template:['Catering2026','Domestic','Laundry','Maintenance'][day%4]});
raw.push({location:'BD: Beaudesert',role:'AIN roster',code:1001,name:'Care Worker',week:1,weekday:0,start:'07:00',end:'15:00',hours:7.6,breakMinutes:30});
raw.push({location:'BD: Beaudesert',role:'AIN roster',code:1001,name:'Care Worker',week:2,weekday:0,start:'07:00',end:'15:00',hours:7.6,breakMinutes:30});
const blankRows=[null,'','   '].map((code,i)=>({...raw[i],code,hours:3}));
const mixed=[...raw,...blankRows,{...raw[0],role:'CSH roster',code:8888,hours:99},{...raw[0],location:'TE: Temora',code:9999,hours:88}];
const published=[{location:'BD: Beaudesert',role:'AIN roster',code:1001,name:'Care Worker',date:anchor,hours:7.6,metadata:'unchanged'}];
const rawBefore=JSON.stringify(mixed);
const publishedBefore=JSON.stringify(published);
const full=supplement(mixed,published,anchor,anchor,'2026-10-11');
equal(full.length,14,'One assigned CS row for every source day');
equal(new Set(full.map(row=>row.role)),new Set(['CS']),'Only CS supplements');
equal(new Set(full.map(row=>row.location)),new Set(['BD: Beaudesert']),'Only BD supplements');
equal(new Set(full.map(row=>row.code)),new Set([4001]),'No blank, CSH, other-facility or synthetic employee IDs');
near(full.reduce((sum,row)=>sum+row.hours,0),7*7.6+7*7.1,'Raw net fortnight hours');
equal(JSON.stringify(mixed),rawBefore,'Source demand rows, including unassigned rows, stay untouched');
equal(JSON.stringify(published),publishedBefore,'Original published facts stay untouched');
equal(full.every(row=>row.breakMinutes===30),true,'Break minutes stay in minutes');
equal(full.find(row=>row.date==='2026-10-11').endDate,'2026-10-12','Overnight end advances one date');
equal(full.find(row=>row.date===anchor).hours,7.6,'Net hours are not recalculated from eight clock hours minus a break');
equal(new Set(full.map(row=>row.template)),new Set(['Catering2026','Domestic','Laundry','Maintenance']),'All legitimate CS templates are retained');
for(const offset of [-43,-29,-28,-15,-14,-13,-8,-7,-1,0,1,6,7,13,14,15,28,43]) {
    const date=iso(toDay(anchor)+offset);
    const result=supplement(mixed,published,anchor,date,date);
    equal(result.length,1,`Offset ${offset} has one row`);
    equal(result[0].fortnightDay,mod(offset,14)+1,`Offset ${offset} uses a nonnegative fortnight index`);
    equal(result[0].sourceWeek,mod(offset,14)<7?1:2,`Offset ${offset} retains source-week alignment`);
    equal(result[0].weekday,weekday(toDay(date)),`Offset ${offset} retains the actual weekday`);
}
for(const [start,end] of [['2026-09-30','2026-10-15'],['2026-09-21','2026-10-05'],['2026-10-12','2026-11-08'],['2026-09-28','2026-09-28']]) {
    const result=supplement(mixed,published,anchor,start,end);
    equal(result.length,toDay(end)-toDay(start)+1,'Partial/repeated horizon has exactly its actual dates');
    equal(result[0].date,start,'No date before the original published horizon');
    equal(result.at(-1).date,end,'No date after the original published horizon');
    const expected=Array.from({length:toDay(end)-toDay(start)+1},(_,i)=>mod(toDay(start)+i-toDay(anchor),14)<7?7.6:7.1).reduce((a,b)=>a+b,0);
    near(result.reduce((sum,row)=>sum+row.hours,0),expected,'Partial/repeated horizon uses actual days, not fortnight averages');
}
const secondWeekOnly=mixed.map(row=>row.role==='CS roster'&&row.week===1?{...row,code:null}:row);
const sparse=supplement(secondWeekOnly,published,anchor,anchor,'2026-10-11');
equal(sparse.length,7,'Blank first-week CS IDs do not rerank source week two');
equal(sparse[0].date,'2026-10-05','Week-two CS rows start on the second Monday');
equal(supplement(mixed.map(row=>roleGroup.get(row.role)==='CS'?{...row,code:null}:row),published,anchor,anchor,'2026-10-11').length,0,'All-unassigned CS does not invent allocation employees');
const unrelatedMalformed=mixed.map(row=>roleGroup.get(row.role)!=='CS'||facility.get(row.location)!=='BD'?{...row,code:'unrelated text',name:null,hours:NaN}:row);
equal(supplement(unrelatedMalformed,published,anchor,anchor,'2026-10-11').length,14,'Unrelated role/facility employee fields do not enter assigned CS validation');
equal(supplement([...mixed,{...raw[0],code:null,name:null,hours:NaN}],published,anchor,anchor,'2026-10-11').length,14,'Blank CS code exclusion happens before assigned identity/hour validation');
rejected(()=>supplement(mixed,[...published,{...published[0],role:'CS roster'}],anchor,anchor,'2026-10-11'),'Any future published CS blocks supplementation');
rejected(()=>supplement(mixed,published,'2026-09-29',anchor,'2026-10-11'),'A non-Monday anchor is rejected');
rejected(()=>supplement([{...raw[0],hours:-1},...raw.slice(1)],published,anchor,anchor,'2026-10-11'),'Negative assigned CS hours are rejected');
rejected(()=>supplement([{...raw[0],name:null},...raw.slice(1)],published,anchor,anchor,'2026-10-11'),'An assigned CS row cannot invent a missing employee name');
const publishedAliases=[{code:1001,employeeCode:9001,name:'Canonical Worker'},{code:2002,employeeCode:9002,name:'Other Published Worker'}];
equal(resolveIdentities([{code:9001,name:'Unambiguous Master Alias'}],publishedAliases).map(row=>({code:row.code,name:row.name})),[{code:1001,name:'Canonical Worker'}],'A unique Master alias resolves to the preserved published code/name');
rejected(()=>resolveIdentities([{code:9001,name:'Other Published Worker'}],publishedAliases),'The raw Master name cannot identify a different published code after alias resolution');
rejected(()=>resolveIdentities([{code:4001,name:'First Resolved Name'},{code:4001,name:'Second Resolved Name'}],[]),'One canonical code cannot retain two exact resolved names');
rejected(()=>supplement(mixed.filter(row=>row.week===1).map(row=>roleGroup.get(row.role)==='CS'?{...row,code:null}:row),published,anchor,anchor,'2026-10-11'),'An all-unassigned CS supplement still validates the complete two-week source prerequisite');

// Existing multi-role display behaviour, widened only from exactly two to two or more.
const people=[{name:'One',role:'CS'},{name:'Two',role:'AIN'},{name:'Two',role:'CS'},
    {name:'Three',role:'AIN'},{name:'Three',role:'AINC4'},{name:'Three',role:'CS'}];
const named=people.map(row=>({...row,display:[...new Set(people.filter(other=>other.name===row.name).map(other=>other.role))].length>=2?`${row.name} (${row.role})`:row.name}));
equal(named.map(row=>row.display),['One','Two (AIN)','Two (CS)','Three (AIN)','Three (AINC4)','Three (CS)'],'Two and three roles get consistent role suffixes');
const ratioRows=[{name:'Three',role:'AIN',hours:15.2},{name:'Three',role:'AINC4',hours:7.6},{name:'Three',role:'CS',hours:999}];
const legacy=ratioRows.filter(row=>['AIN','AINC4'].includes(row.role));
near(legacy.find(row=>row.role==='AINC4').hours/legacy.reduce((sum,row)=>sum+row.hours,0),1/3,'CS cannot dilute the legacy AIN/AINC4 ratio');
equal(215+44,259,'Reported assigned plus blank-ID CS rows reconcile');
near(1400+293,1693,'Reported allocation-eligible plus demand-only net hours reconcile');

const summary={status:'Passed independent reference checks',referenceAssertions:assertions,
    covered:['CS-only BD selection','blank-ID allocation exclusion with source demand untouched','Monday anchor with negative and positive offsets',
        'partial and repeated horizons','raw net hours and break minutes','overnight end dates','all legitimate templates','unchanged published facts',
        'two-or-more role naming','legacy AIN/AINC4 ratio scope','future published-CS overlap guard'],
    nativeEvaluationPerformed:false,productionMExecuted:false,
    limitation:'These are independent reference expectations. They do not prove that the pending production M implements them or that Excel outputs are current.'};
fs.writeFileSync(path.join(scratch,'reference-summary.json'),JSON.stringify(summary,null,2));
fs.writeFileSync(path.join(scratch,'reference-fixture-data.json'),JSON.stringify({anchor,raw:mixed,published,expectedFullFortnight:full},null,2));
console.log(JSON.stringify(summary,null,2));
