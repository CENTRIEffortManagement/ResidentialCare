// Pure reference equivalence, exact M-expression comparison, and syntax parsing only.
// No native M evaluation, Excel access, workbook read, or refresh occurs here.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import crypto from 'node:crypto';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';

const scratch=path.dirname(fileURLToPath(import.meta.url));
const repo=path.resolve(scratch,'../../..');
const sourceRelative='CLIENT/DATExx-Whiddon/UNITS/BD/1. Input/1-AllocationExtracted.xlsx_PowerQuery.m';
const before=fs.readFileSync(path.join(scratch,'source-before.m'),'utf8').replace(/\r\n?/g,'\n');
const after=fs.readFileSync(path.join(repo,sourceRelative),'utf8').replace(/\r\n?/g,'\n');

function definitions(text) {
    const found=[];
    for(const match of text.matchAll(/^shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=/gm)) {
        const start=match.index+match[0].length;
        let quoted=false,lineComment=false,blockComment=false,end=-1;
        for(let i=start;i<text.length;i++) {
            const c=text[i],n=text[i+1];
            if(lineComment){if(c==='\n')lineComment=false;continue;}
            if(blockComment){if(c==='*'&&n==='/'){blockComment=false;i++;}continue;}
            if(quoted){if(c==='"'&&n==='"')i++;else if(c==='"')quoted=false;continue;}
            if(c==='/'&&n==='/'){lineComment=true;i++;continue;}
            if(c==='/'&&n==='*'){blockComment=true;i++;continue;}
            if(c==='"'){quoted=true;continue;}
            if(c===';'){end=i;break;}
        }
        assert.notEqual(end,-1,'Every shared definition terminates');
        const name=match[1].startsWith('#"')?match[1].slice(2,-1).replaceAll('""','"'):match[1];
        found.push({name,expression:text.slice(start,end).trim()});
    }
    return found;
}
const original=definitions(before),current=definitions(after);
const currentMap=new Map(current.map(d=>[d.name,d.expression]));
let sourceContracts=0,referenceAssertions=0;
const contract=(condition,label)=>{sourceContracts++;assert.ok(condition,label);};
const equal=(a,b,label)=>{referenceAssertions++;assert.deepEqual(a,b,label);};
contract(new Set(current.map(d=>d.name)).size===current.length,'No duplicate query names');
contract(JSON.stringify(original.map(d=>d.name))===JSON.stringify(current.map(d=>d.name)),'Every query name and its ordering is preserved');
const changed=[];
for(const definition of original) {
    const expression=currentMap.get(definition.name);
    if(definition.name!=='AllocationExtracted')contract(expression===definition.expression,`Unchanged expression: ${definition.name}`);
    if(expression!==definition.expression)changed.push(definition.name);
}
contract(JSON.stringify(changed)===JSON.stringify(['AllocationExtracted']),'Only AllocationExtracted changes');
contract(currentMap.has('ResDoubleRoles'),'The existing public helper remains available');
const main=currentMap.get('AllocationExtracted');
const tokens=[];
for(const match of main.matchAll(/\/\/[^\r\n]*|\/\*[\s\S]*?\*\/|#"(?:[^"]|"")*"|"(?:[^"]|"")*"|(?<![\w.])[A-Za-z_][A-Za-z0-9_]*(?![\w.])/g)) {
    const token=match[0];
    if(token.startsWith('#"'))tokens.push(token.slice(2,-1).replaceAll('""','"'));
    else if(/^[A-Za-z_]/.test(token))tokens.push(token);
}
contract(tokens.filter(token=>token==='AllocationExtraction').length===1,'The main query references AllocationExtraction once');
contract(!tokens.includes('ResDoubleRoles'),'The main query has no active reference to the shared helper');
contract(/Source\s*=\s*Table\.Buffer\(AllocationExtraction\)/.test(main),'One local source buffer feeds publication and role lookup');
contract(/NameRoles\s*=\s*Table\.SelectColumns\(Source,\s*\{"Name",\s*"Role"\}\)/.test(main),'Role lookup is narrowed to Name and Role from the same source');
contract(/DistinctNameRoles\s*=\s*Table\.Distinct\(NameRoles\)/.test(main),'Role counts use distinct Name/Role pairs');
contract(/Table\.Group\(DistinctNameRoles,\s*\{"Name"\}/.test(main),'Role counts retain the existing Name grain');
contract(/Table\.SelectRows\(WorkerRoleCounts,\s*each\s*\[Count\]\s*>=\s*2\)/.test(main),'Two or more roles retain suffixes');
contract(/Table\.NestedJoin\(\s*Source,\s*\{"Name"\},\s*MultiRoleLookup,/.test(main),'The existing join reads the local role lookup');
const oldMain=original.find(d=>d.name==='AllocationExtracted').expression;
const tailMarker='#"Expanded ResDoubleRoles" =';
contract(main.slice(main.indexOf(tailMarker))===oldMain.slice(oldMain.indexOf(tailMarker)),'Expansion, naming, row projection, final target and column order remain verbatim');

// Two independent constructions of the existing name/role lookup.
const nameKey=(name)=>JSON.stringify(name);
const pairKey=(name,role)=>JSON.stringify([name,role]);
function legacyLookup(rows) {
    const pairGroups=new Map();
    for(const row of rows) {
        const key=pairKey(row.Name,row.Role);
        const group=pairGroups.get(key)??{Name:row.Name,Role:row.Role,Count:0};
        group.Count++;pairGroups.set(key,group);
    }
    const nameGroups=new Map();
    for(const row of pairGroups.values()) {
        const key=nameKey(row.Name);
        const group=nameGroups.get(key)??{Name:row.Name,Count:0};
        group.Count++;nameGroups.set(key,group);
    }
    return [...nameGroups.values()].filter(row=>row.Count>=2).map(row=>({...row,MultiRole:'MultiRole'}));
}
function localLookup(rows) {
    const pairs=[...new Map(rows.map(row=>[pairKey(row.Name,row.Role),{Name:row.Name,Role:row.Role}])).values()];
    const counts=new Map();
    for(const row of pairs)counts.set(nameKey(row.Name),(counts.get(nameKey(row.Name))??0)+1);
    return [...counts].filter(([,count])=>count>=2).map(([name,count])=>({Name:JSON.parse(name),Count:count,MultiRole:'MultiRole'}));
}
function publish(rows,columns,lookup) {
    const outputColumns=columns.filter(column=>column!=='Name').concat('Name');
    const output=rows.map(row=>{
        const matches=lookup.filter(candidate=>candidate.Name===row.Name);
        assert.ok(matches.length<=1,'A grouped name produces at most one join match');
        const name=matches.length?(row.Name===null||row.Role===null?null:`${row.Name} (${row.Role})`):row.Name;
        return Object.fromEntries(outputColumns.map(column=>[column,column==='Name'?name:row[column]]));
    });
    return {columns:outputColumns,rows:output};
}
const baseColumns=['Name','Role','Hours','Code','Order','Extra'];
const rows=(pairs)=>pairs.map(([Name,Role,Hours=7.6],index)=>({Name,Role,Hours,Code:1000+index,Order:index,Extra:{source:`row-${index}`}}));
const cases=[
    {name:'empty',rows:[],names:[]},
    {name:'single role',rows:rows([['Alex','AIN']]),names:['Alex']},
    {name:'duplicate same role',rows:rows([['Alex','AIN',1.25],['Alex','AIN',2],['Alex','AIN',0]]),names:['Alex','Alex','Alex']},
    {name:'two roles AIN CS',rows:rows([['Alex','AIN'],['Alex','CS']]),names:['Alex (AIN)','Alex (CS)']},
    {name:'three roles including CS',rows:rows([['Alex','AIN'],['Alex','AINC4'],['Alex','CS']]),names:['Alex (AIN)','Alex (AINC4)','Alex (CS)']},
    {name:'duplicate one role among two',rows:rows([['Alex','AIN',1.25],['Alex','CS',0],['Alex','AIN',2.5]]),names:['Alex (AIN)','Alex (CS)','Alex (AIN)']},
    {name:'interleaved worker order',rows:rows([['Alex','AIN'],['Beth','RN'],['Alex','CS']]),names:['Alex (AIN)','Beth','Alex (CS)']},
    {name:'null name with two roles',rows:rows([[null,'AIN'],[null,'CS']]),names:[null,null]},
    {name:'null role only',rows:rows([['Alex',null]]),names:['Alex']},
    {name:'null role among two',rows:rows([['Alex',null],['Alex','RN']]),names:[null,'Alex (RN)']},
    {name:'case-sensitive names',rows:rows([['Alex','AIN'],['alex','AIN'],['Alex','CS']]),names:['Alex (AIN)','alex','Alex (CS)']},
    {name:'case-sensitive roles',rows:rows([['Alex','RN'],['Alex','rn']]),names:['Alex (RN)','Alex (rn)']},
    {name:'blank names keep legacy formatting',rows:rows([['','AIN'],['','CS']]),names:[' (AIN)',' (CS)']},
    {name:'null and signed hours remain facts',rows:rows([['Alex','AIN',null],['Alex','CS',-1],['Beth','RN',0.1]]),names:['Alex (AIN)','Alex (CS)','Beth']},
    {name:'spaces are not silently normalised',rows:rows([['Alex','RN'],['Alex ','CS']]),names:['Alex','Alex ']}
];
for(const scenario of cases) {
    const sourceSnapshot=structuredClone(scenario.rows);
    const oldLookup=legacyLookup(scenario.rows),newLookup=localLookup(scenario.rows);
    const legacy=publish(scenario.rows,baseColumns,oldLookup),local=publish(scenario.rows,baseColumns,newLookup);
    equal(local,legacy,`${scenario.name}: complete output equivalence`);
    equal(local.columns,['Role','Hours','Code','Order','Extra','Name'],`${scenario.name}: Name retains the existing final column position`);
    equal(local.rows.map(row=>row.Name),scenario.names,`${scenario.name}: independent expected names`);
    equal(local.rows.map(row=>row.Order),scenario.rows.map(row=>row.Order),`${scenario.name}: output row order`);
    equal(local.rows.map(row=>row.Hours),scenario.rows.map(row=>row.Hours),`${scenario.name}: exact original hours`);
    equal(scenario.rows,sourceSnapshot,`${scenario.name}: input values remain unchanged`);
    equal(newLookup,oldLookup,`${scenario.name}: complete helper lookup equivalence`);
}
for(const columns of [['Hours','Name','Role','Code','Order','Extra'],['Role','Code','Order','Name','Hours','Extra']]) {
    const sample=rows([['Alex','AIN'],['Alex','CS']]);
    equal(publish(sample,columns,localLookup(sample)),publish(sample,columns,legacyLookup(sample)),'Arbitrary original column ordering is preserved');
}

const require=createRequire(import.meta.url);
const bundlePath=process.argv[2];
if(!bundlePath)throw new Error('Pass the installed Microsoft Power Query parser server bundle path.');
let bundle=fs.readFileSync(bundlePath,'utf8');
const entry='var l=t(1422);module.exports=l';
if(!bundle.includes(entry))throw new Error('Unrecognised installed parser bundle.');
bundle=bundle.replace(entry,'module.exports=t(9833)');
const sandbox={module:{exports:{}},require,console,process,Buffer,setTimeout,clearTimeout,setImmediate,clearImmediate};
vm.runInNewContext(bundle,sandbox,{filename:bundlePath});
const pq=sandbox.module.exports;
const result=await pq.TaskUtils.tryLexParse(pq.DefaultSettings,after);
if(pq.TaskUtils.isError(result))throw new Error(JSON.stringify(result.error.innerError??{message:result.error.message}));
const hash=(value)=>crypto.createHash('sha256').update(value).digest('hex').toUpperCase();
const summary={status:'Passed offline equivalence and source checks',source:sourceRelative,sourceSHA256:crypto.createHash('sha256').update(fs.readFileSync(path.join(repo,sourceRelative))).digest('hex').toUpperCase(),
    sourceContracts,referenceCases:cases.length,referenceAssertions,sharedQueryCount:current.length,changedQueryExpressions:changed,
    unchangedOtherQueryExpressions:current.length-changed.length,publicResDoubleRolesRetained:true,
    mainAllocationExtractionReferences:tokens.filter(token=>token==='AllocationExtraction').length,mainExternalResDoubleRolesReferences:tokens.filter(token=>token==='ResDoubleRoles').length,
    syntax:'Passed',parser:'Installed Microsoft Power Query extension parser',nativeEvaluationPerformed:false,productionMExecuted:false,timingMeasured:false,
    limitation:'The arithmetic/reference models and M parser do not execute production M or measure Excel refresh time or memory.'};
fs.writeFileSync(path.join(scratch,'offline-summary.json'),JSON.stringify(summary,null,2));
console.log(JSON.stringify(summary,null,2));
