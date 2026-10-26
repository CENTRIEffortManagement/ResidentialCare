// Exact source preservation, selected source contracts and Microsoft parser only.
// No native M engine, workbook read, connector access, refresh or timing occurs.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import crypto from 'node:crypto';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';

const scratch=path.dirname(fileURLToPath(import.meta.url));
const repo=path.resolve(scratch,'../..');
const sourceRelative='CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity-ShiftAvailability.xlsx_PowerQuery.m';
const sourcePath=path.join(repo,sourceRelative);
const normal=text=>text.replace(/\r\n?/g,'\n');
const before=normal(fs.readFileSync(path.join(scratch,'source-before.m'),'utf8'));
const after=normal(fs.readFileSync(sourcePath,'utf8'));
let assertions=0;
const contract=(condition,label)=>{assertions++;assert.ok(condition,label);};
function definitions(text) {
    const found=[];
    for(const match of text.matchAll(/^shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][A-Za-z0-9_]*)\s*=/gm)) {
        const start=match.index+match[0].length;
        let quoted=false,lineComment=false,blockComment=false,end=-1;
        for(let index=start;index<text.length;index++) {
            const current=text[index],next=text[index+1];
            if(lineComment){if(current==='\n')lineComment=false;continue;}
            if(blockComment){if(current==='*'&&next==='/'){blockComment=false;index++;}continue;}
            if(quoted){if(current==='"'&&next==='"')index++;else if(current==='"')quoted=false;continue;}
            if(current==='/'&&next==='/'){lineComment=true;index++;continue;}
            if(current==='/'&&next==='*'){blockComment=true;index++;continue;}
            if(current==='"'){quoted=true;continue;}
            if(current===';'){end=index;break;}
        }
        assert.notEqual(end,-1,'Every shared definition terminates');
        const name=match[1].startsWith('#"')?match[1].slice(2,-1).replaceAll('""','"'):match[1];
        found.push({name,expression:text.slice(start,end).trim()});
    }
    return found;
}
const original=definitions(before),current=definitions(after);
const old=new Map(original.map(row=>[row.name,row.expression]));
const now=new Map(current.map(row=>[row.name,row.expression]));
const allowed=['IMPORT AllocationExtracted','ReconciledWorkerMembership_Prepare','AvailabilityRoleGroups','AllocationRoleHours_Prepare'];
const additions=['AllocationExtracted_Prepare','CSAllocationWorkers_Prepare','RosterMembership_Prepare'];
contract(original.length===76,'Approved baseline contains 76 definitions');
contract(current.length===79,'Final source contains the expected three new definitions');
contract(current.length===now.size,'No duplicate query names');
contract(JSON.stringify(original.map(row=>row.name))===JSON.stringify(current.filter(row=>old.has(row.name)).map(row=>row.name)),'All original query names and their relative order are retained');
contract(JSON.stringify(current.filter(row=>!old.has(row.name)).map(row=>row.name).sort())===JSON.stringify([...additions].sort()),'Only the three approved queries are added');
const changed=[];
for(const definition of original) {
    if(now.get(definition.name)!==definition.expression)changed.push(definition.name);
    if(!allowed.includes(definition.name))contract(now.get(definition.name)===definition.expression,`Unchanged original expression: ${definition.name}`);
}
contract(JSON.stringify([...changed].sort())===JSON.stringify([...allowed].sort()),'Only the four approved original expressions change');
contract(now.get('ReconciledWorkerMembership_Prepare')===old.get('ReconciledWorkerMembership_Prepare').replace('RosterWorkers_Prepare,','RosterMembership_Prepare,'),'Existing membership, preferred-role, termination and HR rules change only their roster join input');
contract(now.get('AllocationRoleHours_Prepare')===old.get('AllocationRoleHours_Prepare').replace('#"IMPORT AllocationExtracted"','AllocationExtracted_Prepare'),'Existing allocation-hours validation changes only its preparation source');
contract(/List\.Buffer\(\{"RN",\s*"AIN",\s*"EN",\s*"CS"\}\)/.test(now.get('AvailabilityRoleGroups')),'CS is appended to the original approved-role list without CSH');

const landing=now.get('IMPORT AllocationExtracted');
const preparation=now.get('AllocationExtracted_Prepare');
const bridge=now.get('CSAllocationWorkers_Prepare');
const membership=now.get('RosterMembership_Prepare');
contract(!/Table\.(?:SelectRows|SelectColumns|Transform\w*|RenameColumns|AddColumn|RemoveColumns|Group)\s*\(/.test(landing),'Allocation import lands the complete source table without data preparation');
contract(landing.includes('"\\1-AllocationExtracted.xlsx"'),'Existing allocation workbook source is retained');
for(const column of ['Location','Code','Role','Hours','Name','RoleShift','Employee Code'])contract(preparation.includes(`"${column}"`),`Allocation preparation requires ${column}`);
for(const column of ['Location','Code','Role','Name'])contract(preparation.includes(`{"${column}", AvailabilityText, type nullable text}`),`Original ${column} normalization is preserved`);
contract(preparation.includes('#"IMPORT AllocationExtracted"'),'Preparation uses the existing raw allocation landing');
contract(bridge.includes('AllocationExtracted_Prepare'),'Bridge uses the same prepared saved allocation source');
contract(bridge.includes('Table.PromoteHeaders(#"IMPORT Published Roster"'),'Presence guard starts from raw Published Roster before its legacy location filters');
contract(!bridge.includes('PublishedRoster_Prepare'),'Presence guard does not inherit existing published-row filters');
contract(!bridge.includes('RosterWorkers_Prepare'),'Bridge anti-join does not use the agency-filtered population');
contract(bridge.includes('WorkerRoleMappings_Prepare'),'Saved raw RoleShift is checked against existing role mappings');
contract(bridge.includes('AllocationRoleShares'),'Existing validated allocation shares supply role-count evidence');
contract(bridge.includes('"Employee Code"')&&bridge.includes('"Code"'),'Both published/raw employee identity aliases are present');
contract(/Table\.IsEmpty\(\[PublishedByCode\]\)\s+and\s+Table\.IsEmpty\(\[PublishedByAllocationAlias\]\)/.test(bridge),'Existing direct and canonical alias identities are both removed before membership append');
contract(bridge.includes('Table.SelectColumns(AllocationExtracted_Prepare, {"Code", "Employee Code", "Location"})'),'Canonical aliases come from all saved allocation roles');
contract(bridge.includes('PublishedCanonicalCodes')&&/Table\.NestedJoin\(ByCanonicalID,\s*\{"EmployeeID",\s*"Facility-Abbrev"\},\s*PublishedCanonicalCodes/.test(bridge),'Published alias matches block the complete canonical employee/facility key');
contract(bridge.indexOf('AllocationAliases = Table.Distinct')<bridge.indexOf('RelevantAliases = Table.NestedJoin'),'Saved alias history is deduplicated before candidate restriction joins');
const firstJoin=bridge.indexOf('Table.NestedJoin('),firstDistinct=bridge.indexOf('Table.Distinct('),firstSelect=bridge.indexOf('Table.SelectColumns(');
contract(firstSelect>=0&&firstDistinct>firstSelect&&firstJoin>firstDistinct,'Bridge narrows and deduplicates before its first join');
contract(/\[Role\]\s*=\s*"CS"/.test(bridge),'Bridge selects CS rows explicitly');
contract(bridge.includes('Text.EndsWith')&&bridge.includes('" (CS)"'),'Generated suffix handling is explicit and terminal');
contract(/\[RoleCount\]\s*>\s*1/.test(bridge),'Removing the generated suffix requires multiple allocation roles');
contract(membership.includes('Table.Combine')&&membership.includes('RosterWorkers_Prepare')&&membership.includes('CSAllocationWorkers_Prepare'),'New roster membership combines original workers with CS-only additions');

// External source declarations are unchanged. This does not count physical evaluations or reads.
const accessPattern=/\b(?:File\.Contents|Excel\.Workbook|SharePoint\.Tables|Web\.Contents|Folder\.Files)\s*\(/g;
contract((after.match(accessPattern)??[]).length===(before.match(accessPattern)??[]).length,'External file/connector call-site count is unchanged');
contract(!/Master.?Roster|Master Roster/i.test(bridge.replace(/\/\/[^\n]*|\/\*[\s\S]*?\*\//g,'')),'Bridge has no Master Roster source reference');

const bundlePath=process.argv[2];
if(!bundlePath)throw new Error('Pass the installed Microsoft Power Query parser bundle path.');
let bundle=fs.readFileSync(bundlePath,'utf8');
const entry='var l=t(1422);module.exports=l';
if(!bundle.includes(entry))throw new Error('Unrecognized parser bundle.');
bundle=bundle.replace(entry,'module.exports=t(9833)');
const sandbox={module:{exports:{}},require:createRequire(import.meta.url),console,process,Buffer,setTimeout,clearTimeout,setImmediate,clearImmediate};
vm.runInNewContext(bundle,sandbox,{filename:bundlePath});
const pq=sandbox.module.exports;
const result=await pq.TaskUtils.tryLexParse(pq.DefaultSettings,after);
if(pq.TaskUtils.isError(result))throw new Error(JSON.stringify(result.error.innerError??{message:result.error.message}));
const summary={status:'Passed source contracts and Microsoft M syntax parser',source:sourceRelative,sourceSHA256:crypto.createHash('sha256').update(fs.readFileSync(sourcePath)).digest('hex').toUpperCase(),sourceContracts:assertions,originalDefinitions:original.length,currentDefinitions:current.length,changedOriginalQueryExpressions:changed,unchangedOriginalQueryExpressions:original.length-changed.length,newQueryNames:current.filter(row=>!old.has(row.name)).map(row=>row.name),syntax:'Passed',nativeEvaluationPerformed:false,productionMExecuted:false,timingMeasured:false,limitation:'Source contracts and the parser do not execute production M or verify saved workbook data/bindings.'};
fs.writeFileSync(path.join(scratch,'source-summary.json'),JSON.stringify(summary,null,2));
console.log(JSON.stringify(summary,null,2));
