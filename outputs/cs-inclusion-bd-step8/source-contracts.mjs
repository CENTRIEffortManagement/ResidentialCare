// Static production-source contracts and installed Microsoft M parser only.
// This script neither evaluates M nor reads any workbook or external data source.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import crypto from 'node:crypto';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';
const scratch=path.dirname(fileURLToPath(import.meta.url));
const sourceRelative='CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Capacity.xlsx_PowerQuery.m';
const sourcePath=path.resolve(scratch,'../..',sourceRelative);
const normalize=value=>value.replace(/\r\n?/g,'\n');
const before=normalize(fs.readFileSync(path.join(scratch,'source-before.m'),'utf8'));
const after=normalize(fs.readFileSync(sourcePath,'utf8'));
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
function tokens(expression) {
    const result=[];
    for(const match of expression.matchAll(/\/\/[^\n]*|\/\*[\s\S]*?\*\/|#"(?:[^"]|"")*"|"(?:[^"]|"")*"|(?<![\w.])[A-Za-z_][A-Za-z0-9_]*(?![\w.])/g)) {
        const token=match[0];
        if(token.startsWith('#"'))result.push(token.slice(2,-1).replaceAll('""','"'));
        else if(/^[A-Za-z_]/.test(token))result.push(token);
    }
    return result;
}
const original=definitions(before),current=definitions(after);
const old=new Map(original.map(row=>[row.name,row.expression])),now=new Map(current.map(row=>[row.name,row.expression]));
const allowed=['LINKRole','Role1','Role2','Role3','Capacity','AvailabilityDeveloped'];
const additions=['Capacity Enabled Roles Prepare','IMPORT Role A1','IMPORT Role A2','IMPORT Role B','EXTRACT Role Capacity Table','Capacity Stage Prepare','Capacity Role Inputs Prepare','Capacity Final Inputs Prepare','Capacity Reporting Inputs Prepare','Capacity Calendar Prepare','Capacity Calendar Match Prepare'];
contract(original.length===24,'Approved baseline has 24 queries');
contract(current.length===now.size,'No duplicate query names');
contract(current.length===original.length+additions.length,'Only the expected eleven queries are added');
contract(JSON.stringify(original.map(row=>row.name))===JSON.stringify(current.filter(row=>old.has(row.name)).map(row=>row.name)),'Original query names and relative order are retained');
contract(JSON.stringify(current.filter(row=>!old.has(row.name)).map(row=>row.name).sort())===JSON.stringify([...additions].sort()),'Only approved new helpers and landings are added');
const changed=[];
for(const row of original) {
    if(row.expression!==now.get(row.name))changed.push(row.name);
    if(!allowed.includes(row.name))contract(row.expression===now.get(row.name),`Unchanged original expression: ${row.name}`);
}
contract(JSON.stringify([...changed].sort())===JSON.stringify([...allowed].sort()),'Exactly the six approved original expressions change');
for(const [name,label] of [['Role1','AIN'],['Role2','RN'],['Role3','EN']])contract(now.get(name)===`"${label}"`,`${name} keeps its stable legacy nursing label`);
const raw=now.get('LINKRole');
contract(raw.includes('SharePoint.Tables')&&raw.includes('bad3beb8-7064-4e11-9edd-d62dac5d702d'),'Original Role Analysis list choice is preserved');
contract(!/Table\.(?:SelectRows|SelectColumns|Transform\w*|AddIndexColumn|AddColumn|RemoveColumns|RenameColumns)\s*\(/.test(raw),'LINKRole lands raw list items without business preparation');
const enabled=now.get('Capacity Enabled Roles Prepare');
contract(enabled.includes('LINKRole')&&enabled.includes('Effort Management Analysis'),'Enabled role labels come from the existing list and flag');
contract(enabled.includes('Table.Sort')&&enabled.includes('Table.Group')&&enabled.includes('[Rows] <> 1'),'Enabled labels are sorted and duplicate labels are rejected');
contract(!/"(?:AIN|RN|EN|CS|CSH|AINC4)"/.test(enabled),'Enabled roles have no hard-coded role-label whitelist');
for(const [name,file] of [['IMPORT Role A1','CapacityDistrib(A.1)-shifts.xlsx'],['IMPORT Role A2','CapacityDistrib(A.2)-shifts.xlsx'],['IMPORT Role B','CapacityDistrib(B)-shifts.xlsx']]) {
    const expression=now.get(name);
    contract(expression.includes(file)&&expression.includes('Folder'),'Approved role workbook and existing unit folder are preserved: '+name);
    contract(expression.includes('Binary.Buffer')&&expression.includes('Excel.Workbook')&&expression.includes('File.Contents'),'The role landing buffers its source binary and returns workbook navigation: '+name);
    contract(!/Table\.(?:SelectRows|SelectColumns|Transform\w*|AddColumn|RemoveColumns|RenameColumns|Group)\s*\(/.test(expression),'Role import does not transform source data: '+name);
    contract(!/\botherwise\b/.test(expression),'Missing role workbooks do not silently fall back: '+name);
}
const extract=now.get('EXTRACT Role Capacity Table');
contract(extract.includes('Kind')&&extract.includes('"Table"')&&extract.includes('[Data]'),'Extraction navigates the exact saved table kind');
contract(!/\botherwise\b|Table\.(?:SelectRows|SelectColumns|Transform\w*|AddColumn|RemoveColumns|RenameColumns|Group)\s*\(/.test(extract),'Saved table extraction has no transformations or fallback');
const roleInputs=now.get('Capacity Role Inputs Prepare');
for(const table of ['ResPeriodAvailabilityTABLE','ResPeriodAvailabilityCapped_C__TABLE','C__TABLE','C___TABLE_B'])contract(roleInputs.includes(`"${table}"`),'Existing saved table Item is preserved: '+table);
contract(tokens(roleInputs).filter(token=>token==='IMPORT Role B').length===1,'One local B landing supplies both B stages within one role-helper evaluation');
contract(tokens(roleInputs).includes('IMPORT Role A1')&&tokens(roleInputs).includes('IMPORT Role A2'),'The reporting branch uses genuine A1 and A2 landings');
const stage=now.get('Capacity Stage Prepare');
contract(stage.includes('Int64.Type')&&stage.includes('Table.TransformColumnTypes'),'Legacy count/key conversions remain on the preparation side');
contract(stage.includes('AvailabilityType')&&stage.includes('"Availability"'),'Stage preparation supplies the existing value and stage fields');
contract(!stage.includes('Table.ReplaceValue'),'Calculation values are not changed from zero to null by stage preparation');
contract(stage.includes('then {"ResAvailability", "ResMaxAvail"} else {}'),'C### resource totals remain required even for an empty publication');
contract(stage.includes('[Role] <> RoleName')&&stage.includes('MissingKeys')&&stage.includes('DuplicateKeys'),'Wrong role ownership and missing/duplicate resource-period keys block publication');
contract(stage.indexOf('ScopedColumns = if Reporting')<stage.indexOf('Snapshot = Table.Buffer'),'Reporting narrows diagnostic width before buffering');
contract(roleInputs.includes('Result = if not AllStages then Final else'),'Final-only evaluation does not enumerate A1/A2 reporting tables');
const finalInputs=now.get('Capacity Final Inputs Prepare'),reportInputs=now.get('Capacity Reporting Inputs Prepare');
contract(finalInputs.includes('Capacity Enabled Roles Prepare')&&finalInputs.includes('Capacity Role Inputs Prepare')&&finalInputs.includes('false'),'Final inputs enumerate every enabled role through the final-only branch');
contract(reportInputs.includes('Capacity Enabled Roles Prepare')&&reportInputs.includes('Capacity Role Inputs Prepare')&&reportInputs.includes('true'),'Reporting inputs enumerate every enabled role through all stages');
function closure(name,seen=new Set()) {
    if(seen.has(name))return seen;seen.add(name);
    for(const token of tokens(now.get(name)??''))if(now.has(token))closure(token,seen);
    return seen;
}
for(const name of ['Capacity','AvailabilityDeveloped']) {
    const dependencies=closure(name);
    contract(!['Role1','Role2','Role3'].some(role=>dependencies.has(role)),`${name} has no positional role dependency`);
    contract(!original.filter(row=>row.name.startsWith('IMPORT ')).some(row=>dependencies.has(row.name)),`${name} has no active legacy positional import dependency`);
}
const report=now.get('AvailabilityDeveloped'),oldReport=old.get('AvailabilityDeveloped');
contract(report.slice(report.indexOf('#"Added Custom" ='))===oldReport.slice(oldReport.indexOf('#"Added Custom" =')),'Facility, zero-to-null reporting, metadata removal and exact eight-column projection remain verbatim');
contract(now.get('Capacity').includes('Capacity Final Inputs Prepare')&&now.get('Capacity').includes('Capacity Calendar Match Prepare'),'Main Capacity reads generic final inputs and validated calendar matching');
const calendar=now.get('Capacity Calendar Prepare'),match=now.get('Capacity Calendar Match Prepare');
contract(calendar.includes('Scoped = Table.SelectRows')&&calendar.includes('List.Contains(Roles, [RolesList])'),'Calendar validation is scoped to enabled labels');
contract(calendar.includes('Table.Group(Scoped, {"Period", "RolesList"}')&&calendar.includes('[Rows] <> 1'),'Calendar keys are unique by role and period');
contract(match.includes('Table.RowCount([CapacityDate]) <> 1')&&match.includes('else error'),'Every consumed capacity cell requires exactly one date match');
contract(match.includes('Snapshot = Table.Buffer(Joined)'),'Calendar matching buffers one local joined snapshot');
contract(match.includes('Invalid = Table.SelectRows(Snapshot,')&&match.includes('Validated = if Table.IsEmpty(Invalid) then Snapshot'),'Calendar validation and output consume the same local snapshot');
contract(reportInputs.includes('Capacity Calendar Match Prepare'),'Reporting cells also pass the required calendar match gate');
const parserPath=process.argv[2];
if(!parserPath)throw new Error('Pass the installed Microsoft Power Query parser bundle path.');
let bundle=fs.readFileSync(parserPath,'utf8');
const entry='var l=t(1422);module.exports=l';
if(!bundle.includes(entry))throw new Error('Unrecognized parser bundle.');
bundle=bundle.replace(entry,'module.exports=t(9833)');
const sandbox={module:{exports:{}},require:createRequire(import.meta.url),console,process,Buffer,setTimeout,clearTimeout,setImmediate,clearImmediate};
vm.runInNewContext(bundle,sandbox,{filename:parserPath});
const pq=sandbox.module.exports,result=await pq.TaskUtils.tryLexParse(pq.DefaultSettings,after);
if(pq.TaskUtils.isError(result))throw new Error(JSON.stringify(result.error.innerError??{message:result.error.message}));
const summary={status:'Passed source contracts and Microsoft M parser',source:sourceRelative,sourceSHA256:crypto.createHash('sha256').update(fs.readFileSync(sourcePath)).digest('hex').toUpperCase(),sourceContracts:assertions,baselineQueries:original.length,currentQueries:current.length,changedOriginalQueryExpressions:changed,unchangedOriginalQueryExpressions:original.length-changed.length,newQueryNames:current.filter(row=>!old.has(row.name)).map(row=>row.name),syntax:'Passed',nativeEvaluationPerformed:false,productionMExecuted:false,timingMeasured:false,limitation:'Parser and source contracts do not execute M, measure physical file reads or verify saved workbook data/bindings.'};
fs.writeFileSync(path.join(scratch,'source-summary.json'),JSON.stringify(summary,null,2));
console.log(JSON.stringify(summary,null,2));
