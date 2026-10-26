// Offline reference arithmetic, source contracts, and Microsoft M parsing only.
// This script never invokes PQTest, Excel, SharePoint, or an M evaluator.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import crypto from 'node:crypto';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';

const scratch=path.dirname(fileURLToPath(import.meta.url));
const repo=path.resolve(scratch,'../..');
const sourceRelative='CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/2-DemandExtract.xlsx_PowerQuery.m';
const source=fs.readFileSync(path.join(repo,sourceRelative),'utf8');
const fixture=fs.readFileSync(path.join(scratch,'fixtures.pq'),'utf8');
const generated=fs.readFileSync(path.join(scratch,'generated-extension.pq'),'utf8');
const require=createRequire(import.meta.url);
let sourceContracts=0,referenceAssertions=0;
const contract=(ok,message)=>{sourceContracts++;assert.ok(ok,message);};
const near=(a,b,message)=>{referenceAssertions++;assert.ok(Math.abs(a-b)<=1e-7,`${message}: ${a} vs ${b}`);};
const equal=(a,b,message)=>{referenceAssertions++;assert.deepEqual(a,b,message);};
const definitionStarts=[...source.matchAll(/^shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=/gm)];
const definitions=new Map(definitionStarts.map((m,i)=>{
    const name=m[1].startsWith('#"')?m[1].slice(2,-1).replaceAll('""','"'):m[1];
    return [name,source.slice(m.index,definitionStarts[i+1]?.index??source.length)];
}));
const names=[...definitions.keys()];
contract(names.length===definitionStarts.length,'Shared query names must be unique');
const beforePath=path.join(scratch,'legacy-contract-before.json');
if(fs.existsSync(beforePath)) {
    const before=JSON.parse(fs.readFileSync(beforePath,'utf8'));
    contract(JSON.stringify(names)===JSON.stringify(before.queryNames),'All original query names and order remain stable');
}
for(const name of ['ShiftUnitDemandHRS','DemandExtraction_CHECK','DemandExtraction_RECONCILIATION',
    'EXTRACT Distributed Demand FTE','EXTRACT Distributed Demand FTE Fortnight','EXTRACT Distributed Demand FTE Checks',
    'Distributed FTE Rows Prepare','Distributed FTE Source Cells','Distributed FTE Prepare','Demand Roles',
    'Demand Calendar CHECK','Demand Calendar Prepare','Permutation DateTimeRoleShift','Demand Reconciliation Targets']) {
    contract(definitions.has(name),`Required public or compatibility query retained: ${name}`);
}
contract(definitions.get('EXTRACT Distributed Demand FTE').includes('RoleDemand_TABLE'),'Use the generic handover table');
contract(definitions.get('EXTRACT Distributed Demand FTE Checks').includes('RoleDemand_CHECK'),'Use the generic handover checks');
contract(!/Item\s*=\s*"MinuteWorkersFTE/.test(source),'No legacy publication navigation remains');
const published=['Date','Day','Shift','Period','Role','StartTime','EndTime','Unit','Facility','DemandFTE','DemandHRS','DurationOfShifts'];
const publication=definitions.get('ShiftUnitDemandHRS');
for(const column of published) contract(publication.includes(`"${column}"`),`Published column retained: ${column}`);
contract(/\[DemandHRS\]\s*\/\s*\[DurationOfShifts\]/.test(definitions.get('Demand Extraction Prepare')),'Attendance derives from authoritative hours and actual duration');
contract(definitions.get('Permutation DateTimeRoleShift').includes('Date.AddDays([Date], 1)'),'Night shifts retain next-day endpoints');
contract(definitions.get('Demand Calendar Prepare').includes('Demand Calendar CHECK'),'Settings role coverage remains an enforced dependency');
contract(definitions.get('ShiftUnitDemandHRS').includes('DemandExtraction_CHECK'),'Publication remains gated by named checks');
contract(definitions.get('DemandExtraction_CHECK').includes('DemandExtraction_RECONCILIATION'),'Reconciliation remains part of the publication gate');
contract(!/\b(?:File|Folder|Web|Excel|SharePoint|OData|Sql|OleDb|Odbc|Access|AzureStorage)\.[A-Za-z]+\s*\(|\bValue\.NativeQuery\s*\(/i.test(generated),'Native fixture package contains no external connector calls');

// Independent arithmetic uses a small explicit input contract, without loading M functions.
const roles=['RN','AIN','EN','CS'];
const base={RN:2,AIN:3,EN:4,CS:5};
const percentages={RN:1,AIN:1,EN:0.5};
const shifts=[{name:'AM',start:6,duration:8.25},{name:'PM',start:14,duration:7.25},{name:'NIGHT',start:21,duration:9}];
const pattern=[];
for(const role of roles)for(let day=1;day<=14;day++)for(let shift=0;shift<3;shift++) {
    const week=day<=7?1:2;
    const hours=role==='CS'&&day===3&&shift===1?0:base[role]*(week===1?1:2)+(day-1)%7/10+(shift+1)/100;
    pattern.push({role,day,shift,weekNo:week===1?17:41,hours,productive:role==='CS'?null:hours*percentages[role]});
}
equal(pattern.length,168,'Four source role patterns each have 42 cells');
equal(new Set(pattern.map(r=>`${r.role}|${r.day}|${r.shift}`)).size,168,'Source role/day/shift cells are unique');
equal([...new Set(pattern.map(r=>r.weekNo))],[17,41],'Original source-week keys remain available');
const plan=(days)=>{
    const output=[];
    for(let day=0;day<days;day++)for(let shift=0;shift<3;shift++)for(const role of roles) {
        const cell=pattern.find(r=>r.role===role&&r.day===day%14+1&&r.shift===shift);
        const timing=shifts[shift];
        output.push({...cell,planningDay:day+1,period:day*3+shift+1,cycle:Math.floor(day/14)+1,
            startHour:day*24+timing.start,endHour:day*24+timing.start+timing.duration,duration:timing.duration,
            attendance:cell.hours/timing.duration});
    }
    return output;
};
for(const horizon of [1,7,14,15,21,28,35]) {
    const rows=plan(horizon);
    equal(rows.length,horizon*3*4,`${horizon}-day output cardinality`);
    equal(new Set(rows.map(r=>`${r.role}|${r.period}`)).size,rows.length,`${horizon}-day role/period grain`);
    for(const row of rows) {
        near(row.attendance*row.duration,row.hours,'Attendance/hour round trip');
        near(row.endHour-row.startHour,row.duration,'Exact timestamp span');
        if(row.shift===2) equal(Math.floor(row.endHour/24),Math.floor(row.startHour/24)+1,'Night crosses one calendar date');
    }
    for(const role of roles) {
        const sourceRole=pattern.filter(r=>r.role===role);
        const full=Math.floor(horizon/14),remaining=horizon%14;
        const independent=full*sourceRole.reduce((sum,r)=>sum+r.hours,0)+sourceRole.filter(r=>r.day<=remaining).reduce((sum,r)=>sum+r.hours,0);
        near(rows.filter(r=>r.role===role).reduce((sum,r)=>sum+r.hours,0),independent,`${role} ${horizon}-day total with partial fortnight`);
    }
}
const cs=pattern.filter(r=>r.role==='CS');
equal(cs.every(r=>r.productive===null),true,'CS productive hours remain null');
equal(pattern.filter(r=>r.role!=='CS').every(r=>Number.isFinite(r.productive)),true,'Care productive hours remain numeric');
equal(cs.find(r=>r.day===3&&r.shift===1).hours,0,'A proven source zero remains zero');
assert.notEqual(cs.find(r=>r.day===1&&r.shift===0).hours,cs.find(r=>r.day===8&&r.shift===0).hours);referenceAssertions++;
const partial=plan(15).filter(r=>r.cycle===2);
equal(partial.length,12,'Final one-day partial fortnight has exactly 12 role/shift cells');
for(const role of roles)near(partial.filter(r=>r.role===role).reduce((s,r)=>s+r.hours,0),pattern.filter(r=>r.role===role&&r.day===1).reduce((s,r)=>s+r.hours,0),'Partial fortnight uses actual day one');

const bundlePath=process.argv[2];
if(!bundlePath)throw new Error('Pass the installed Microsoft Power Query parser server bundle path.');
let bundle=fs.readFileSync(bundlePath,'utf8');
const entry='var l=t(1422);module.exports=l';
if(!bundle.includes(entry))throw new Error('Unrecognised installed Microsoft parser bundle.');
bundle=bundle.replace(entry,'module.exports=t(9833)');
const sandbox={module:{exports:{}},require,console,process,Buffer,setTimeout,clearTimeout,setImmediate,clearImmediate};
vm.runInNewContext(bundle,sandbox,{filename:bundlePath});
const pq=sandbox.module.exports;
const parsed=[];
for(const [label,text] of [[sourceRelative,source],['fixtures.pq',fixture],['generated-extension.pq',generated]]) {
    const result=await pq.TaskUtils.tryLexParse(pq.DefaultSettings,text);
    if(pq.TaskUtils.isError(result))throw new Error(`${label}: ${JSON.stringify(result.error.innerError??{message:result.error.message})}`);
    parsed.push({source:label,syntax:'Passed',stage:result.stage});
}
const summary={status:'Passed offline checks',source:sourceRelative,sourceSHA256:crypto.createHash('sha256').update(source).digest('hex').toUpperCase(),
    sourceContracts,referenceArithmeticAssertions:referenceAssertions,parser:'Installed Microsoft Power Query extension parser',parsed,
    nativeFixturesPrepared:(fixture.match(/^\s*Check\("/gm)??[]).length,nativeFixturesExecuted:0,nativeEvaluationPerformed:false,
    limitation:'Reference arithmetic and structural/syntax checks do not execute production M or verify workbook bindings, live data, or refresh.'};
fs.writeFileSync(path.join(scratch,'offline-summary.json'),JSON.stringify(summary,null,2));
console.log(JSON.stringify(summary,null,2));
