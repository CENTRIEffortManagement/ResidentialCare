// Synthetic reference arithmetic and publication contracts. No M runtime or workbook access.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const scratch=path.dirname(fileURLToPath(import.meta.url));
const reportColumns=['Role','Resource','Period','Availability','AvailabilityType','Facility','ResAvailability','ResMaxAvail'];
const matrixKeys=['Role','Resource','Period','Facility'];
const stageSpecs=[
    {Stage:'OriginalAvailability',File:'A1',Table:'ResPeriodAvailabilityTABLE',Value:'Availability'},
    {Stage:'C#',File:'A2',Table:'ResPeriodAvailabilityCapped_C__TABLE',Value:'AvailabilityCapped'},
    {Stage:'C##',File:'B',Table:'C__TABLE',Value:'C##'},
    {Stage:'C###',File:'B',Table:'C___TABLE_B',Value:'C###'}
];
let assertions=0;const cases=[];
const equal=(actual,expected,label)=>{assertions++;assert.deepEqual(actual,expected,label);};
const throws=(fn,pattern,label)=>{assertions++;assert.throws(fn,pattern,label);};
function scenario(name,fn){fn();cases.push(name);}
const sum=values=>values.every(value=>value===null)?null:values.reduce((total,value)=>total+(value??0),0);
function enabledRoles(analysis){
    const roles=analysis.filter(row=>row.Enabled===true).map(row=>typeof row.Role==='string'?row.Role.trim():null);
    if(!roles.length||roles.some(role=>role===null||role===''||role==='.'||role==='..'||/[\\/]/.test(role))||new Set(roles.map(role=>role.toUpperCase())).size!==roles.length)throw new Error('Invalid or duplicate enabled role labels');
    return roles.sort();
}
function table(columns,rows){return {columns,rows};}
function fixture(roles=['AIN','EN','RN','CS'],emptyRoles=[]) {
    const books={};
    roles.forEach((role,index)=>{
        const base=[{Resource:1,Period:1,Role:role},{Resource:1,Period:2,Role:role}];
        const rows=emptyRoles.includes(role)?[]:base;
        books[role]={
            A1:[{Item:'ResPeriodAvailabilityTABLE',Kind:'Table',Data:table(['Role','Resource','Period','Availability','OriginalAvailability','SpareIssueReason'],rows.map((row,i)=>({...row,Availability:i,OriginalAvailability:i,SpareIssueReason:`diagnostic-${role}-${i}`})))}],
            A2:[{Item:'ResPeriodAvailabilityCapped_C__TABLE',Kind:'Table',Data:table(['Role','Resource','Period','AvailabilityCapped','RosteredPeriodStatus','A2Audit'],rows.map((row,i)=>({...row,AvailabilityCapped:i,RosteredPeriodStatus:null,A2Audit:i%2===0})))}],
            B:[
                {Item:'C__TABLE',Kind:'Table',Data:table(['Resource','Period','Role','C##'],rows.map((row,i)=>({...row,'C##':i})))},
                {Item:'C___TABLE_B',Kind:'Table',Data:table(['Resource','Period','Role','C###','ResAvailability','ResMaxAvail','FinalAudit'],rows.map((row,i)=>({...row,'C###':i===0?0:index+1,ResAvailability:10+index,ResMaxAvail:20+index,FinalAudit:`final-${role}-${i}`})))}
            ]
        };
    });
    const analysis=roles.map(Role=>({Role,Enabled:true})).concat({Role:'CSH',Enabled:false});
    const calendar=roles.flatMap(Role=>[1,2].map(Period=>({Role,Period,Date:`2026-10-${25+Period}`})));
    return {analysis,books,calendar};
}
function publish(inputs) {
    const roles=enabledRoles(inputs.analysis),reads=[];
    const stages=[];const finalTables=[];
    for(const role of roles) {
        const snapshots={};
        for(const file of ['A1','A2','B']) {
            const navigation=inputs.books[role]?.[file];
            if(!navigation)throw new Error(`Missing required file ${role}/${file}`);
            reads.push(`${role}/${file}`);snapshots[file]=navigation;
        }
        for(const spec of stageSpecs) {
            const matches=snapshots[spec.File].filter(row=>row.Kind==='Table'&&row.Item===spec.Table);
            if(matches.length!==1)throw new Error(`Missing or ambiguous required table ${role}/${spec.Table}`);
            const source=matches[0].Data;
            const required=['Role','Resource','Period',spec.Value].concat(spec.Stage==='C###'?['ResAvailability','ResMaxAvail']:[]);
            if(required.some(column=>!source.columns.includes(column)))throw new Error(`Missing required column ${role}/${spec.Stage}`);
            const preparedColumns=source.columns.map(column=>column===spec.Value?'Availability':column).concat('AvailabilityType');
            const cellKeys=new Set();
            const prepared=source.rows.map(row=>{
                if(row.Role!==role)throw new Error(`Unexpected row role in ${role}/${spec.Stage}`);
                if(row.Resource===null||row.Period===null)throw new Error('Missing resource/period key');
                const cellKey=JSON.stringify([row.Resource,row.Period]);
                if(cellKeys.has(cellKey))throw new Error('Duplicate resource/period key');
                cellKeys.add(cellKey);
                return Object.fromEntries(preparedColumns.map(column=>[column,column==='AvailabilityType'?spec.Stage:column==='Availability'?row[spec.Value]:row[column]]));
            });
            if(spec.Stage==='C###')finalTables.push({columns:preparedColumns,rows:prepared});
            stages.push(...prepared.map(row=>Object.fromEntries(reportColumns.map(column=>[column,column==='Facility'?'BD':column==='Availability'?(row.Availability===0?null:row.Availability):(row[column]??null)]))));
        }
    }
    const calendarKeys=new Map();
    for(const row of inputs.calendar.filter(row=>roles.includes(row.Role))) {
        const key=JSON.stringify([row.Role,row.Period]);
        if(calendarKeys.has(key))throw new Error('Duplicate Role/Period calendar key');
        calendarKeys.set(key,row.Date);
    }
    const finalRows=finalTables.flatMap(source=>source.rows).map(row=>{
        const calendarKey=JSON.stringify([row.Role,row.Period]);
        if(!calendarKeys.has(calendarKey))throw new Error('Missing Role/Period calendar key');
        return {...row,Date:calendarKeys.get(calendarKey)};
    }).sort((a,b)=>a.Role.localeCompare(b.Role)||a.Resource-b.Resource||a.Period-b.Period);
    const finalColumns=[...new Set(finalTables.flatMap(source=>source.columns))].concat('Date');
    const stageNames=[...new Set(stages.map(row=>row.AvailabilityType))];
    const groups=new Map();
    for(const row of stages) {
        const key=JSON.stringify(matrixKeys.map(column=>row[column]));
        const group=groups.get(key)??{keys:Object.fromEntries(matrixKeys.map(column=>[column,row[column]])),values:new Map()};
        const values=group.values.get(row.AvailabilityType)??[];values.push(row.Availability);group.values.set(row.AvailabilityType,values);groups.set(key,group);
    }
    const matrixRows=[...groups.values()].map(group=>({...group.keys,...Object.fromEntries(stageNames.map(stage=>[stage,group.values.has(stage)?sum(group.values.get(stage)):null]))}));
    return {Roles:roles,LogicalReads:reads,Capacity:{columns:finalColumns,rows:finalRows},Report:{columns:reportColumns,rows:stages},Matrix:{columns:matrixKeys.concat(stageNames),rows:matrixRows}};
}
scenario('all four enabled role labels supply all four publications',()=>{
    const result=publish(fixture());
    equal(result.Roles,['AIN','CS','EN','RN'],'Role identity is independent of input positions');
    equal(result.Capacity.rows.length,8,'Two final rows per role');
    equal(result.Report.rows.length,32,'Four stages times two rows times four roles');
    equal([...new Set(result.Report.rows.map(row=>row.AvailabilityType))],stageSpecs.map(row=>row.Stage),'All four stage names are retained');
    equal(result.LogicalReads.length,12,'Reference model opens three role workbooks per role');
    equal(new Set(result.LogicalReads).size,12,'The B snapshot supplies both B stages');
});
scenario('SharePoint row order cannot omit or duplicate roles',()=>{
    const input=fixture(),before=publish(input);input.analysis.reverse();
    equal(publish(input),before,'Complete outputs are unchanged after role-list reversal');
    input.analysis=[input.analysis[2],input.analysis[4],input.analysis[1],input.analysis[3],input.analysis[0]];
    equal(publish(input),before,'Complete outputs are unchanged after nontrivial role-list permutation');
});
scenario('disabled CS and CSH need no corresponding source files',()=>{
    const input=fixture();input.analysis.find(row=>row.Role==='CS').Enabled=false;delete input.books.CS;
    const result=publish(input);
    equal(result.Roles,['AIN','EN','RN'],'Only enabled labels are published');
    equal(result.LogicalReads.length,9,'Disabled roles are never opened by the reference model');
    equal(result.Report.rows.some(row=>row.Role==='CS'||row.Role==='CSH'),false,'No disabled role rows');
});
scenario('one enabled role needs only its three workbooks',()=>{
    const result=publish(fixture(['CS']));
    equal(result.Roles,['CS'],'No minimum of three positional roles');
    equal([result.LogicalReads.length,result.Capacity.rows.length,result.Report.rows.length],[3,2,8],'One complete role pipeline');
});
scenario('complete configured role labels remain intact',()=>{
    const result=publish(fixture(['AIN','CS Extended']));
    equal(result.Roles,['AIN','CS Extended'],'No partial role-key derivation');
    equal(result.Capacity.rows.filter(row=>row.Role==='CS Extended').length,2,'Exact configured label is used for source identity and calendar joins');
});
scenario('duplicate and invalid enabled labels fail before source access',()=>{
    const duplicate=fixture();duplicate.analysis.push({Role:'cs',Enabled:true});
    throws(()=>publish(duplicate),/duplicate enabled role labels/,'Case variants cannot double count the same folder');
    const blank=fixture();blank.analysis.push({Role:' ',Enabled:true});
    throws(()=>publish(blank),/Invalid/,'Blank enabled labels are rejected');
});
scenario('same resource and period numbers across roles match the correct dates once',()=>{
    const result=publish(fixture());
    equal(result.Capacity.rows.map(row=>[row.Role,row.Period,row.Date]),['AIN','CS','EN','RN'].flatMap(role=>[[role,1,'2026-10-26'],[role,2,'2026-10-27']]),'Calendar match uses Role and Period');
    equal(result.Capacity.rows.reduce((total,row)=>total+(row.Availability??0),0),10,'Final capacity totals are conserved by the date join');
});
scenario('main capacity retains B columns, extras and column order',()=>{
    const result=publish(fixture());
    equal(result.Capacity.columns,['Resource','Period','Role','Availability','ResAvailability','ResMaxAvail','FinalAudit','AvailabilityType','Date'],'B value is renamed in place; extras survive; stage and date are added');
    equal(result.Capacity.rows[0].FinalAudit,'final-AIN-0','Extra B evidence is preserved on Capacity');
    equal(result.Capacity.rows[0].AvailabilityType,'C###','Main publication uses only final capacity');
});
scenario('zero to null is restricted to reporting Availability',()=>{
    const result=publish(fixture());
    equal(result.Capacity.rows.filter(row=>row.Period===1).map(row=>row.Availability),[0,0,0,0],'Main Capacity keeps zero');
    equal(result.Report.rows.filter(row=>row.Period===1).every(row=>row.Availability===null),true,'Reporting hides zero Availability');
    equal(result.Report.rows.find(row=>row.Role==='AIN'&&row.AvailabilityType==='C###').ResAvailability,10,'Resource audit units are unchanged');
    equal(result.Report.rows.find(row=>row.Role==='AIN'&&row.AvailabilityType==='C###').ResMaxAvail,20,'Resource maximum units are unchanged');
});
scenario('nullable publication values remain nullable',()=>{
    const input=fixture();input.books.AIN.B[1].Data.rows[1]['C###']=null;input.books.AIN.B[1].Data.rows[1].ResAvailability=null;
    const result=publish(input);
    equal(result.Capacity.rows.find(row=>row.Role==='AIN'&&row.Period===2).Availability,null,'No null-to-zero fabrication in calculation output');
    equal(result.Report.rows.find(row=>row.Role==='AIN'&&row.Period===2&&row.AvailabilityType==='C###').ResAvailability,null,'Nullable audit value is preserved');
});
scenario('diagnostic columns cannot become reporting or matrix keys',()=>{
    const result=publish(fixture());
    equal(result.Report.columns,reportColumns,'Exact eight-column reporting contract');
    equal(result.Matrix.columns,matrixKeys.concat(stageSpecs.map(row=>row.Stage)),'Only four business keys and four stage measures');
    equal(result.Matrix.rows.length,8,'Stage-specific diagnostics do not split matrix rows');
    equal(result.Matrix.rows.some(row=>'FinalAudit'in row||'A2Audit'in row||'SpareIssueReason'in row),false,'No diagnostic fields leak into the pivot');
});
scenario('matrix null sums retain the existing display convention',()=>{
    const result=publish(fixture());
    equal(result.Matrix.rows.find(row=>row.Role==='CS'&&row.Period===1),{Role:'CS',Resource:1,Period:1,Facility:'BD',OriginalAvailability:null,'C#':null,'C##':null,'C###':null},'All-zero measures remain null on the reporting matrix');
});
scenario('an enabled role with valid empty source tables contributes zero rows',()=>{
    const result=publish(fixture(['AIN','EN','RN','CS'],['CS']));
    equal(result.LogicalReads.length,12,'Enabled empty role still requires its genuine publications');
    equal(result.Capacity.rows.length,6,'No fabricated CS capacity rows');
    equal(result.Report.rows.length,24,'No fabricated CS reporting rows');
});
scenario('all enabled sources may be valid empty tables',()=>{
    const result=publish(fixture(['AIN','EN','RN','CS'],['AIN','EN','RN','CS']));
    equal(result.Capacity.rows,[],'Zero final rows');equal(result.Report.rows,[],'Zero report rows');equal(result.Report.columns,reportColumns,'Report schema remains present');
});
scenario('missing enabled role workbook fails without role substitution',()=>{
    const input=fixture();delete input.books.CS.A2;
    throws(()=>publish(input),/Missing required file CS\/A2/,'No fallback to AIN or another stage');
});
scenario('missing required saved table fails for every stage',()=>{
    for(const spec of stageSpecs) {
        const input=fixture();input.books.CS[spec.File]=input.books.CS[spec.File].filter(row=>row.Item!==spec.Table);
        throws(()=>publish(input),/Missing or ambiguous required table/,'Required saved table is mandatory: '+spec.Stage);
    }
});
scenario('table navigation requires the table kind and unique name',()=>{
    const wrongKind=fixture();wrongKind.books.CS.B[1].Kind='Sheet';
    throws(()=>publish(wrongKind),/required table/,'A similarly named sheet is not a table fallback');
    const duplicate=fixture();duplicate.books.CS.B.push(structuredClone(duplicate.books.CS.B[1]));
    throws(()=>publish(duplicate),/ambiguous/,'Ambiguous table navigation does not choose a first match');
});
scenario('required columns are checked even on empty publications',()=>{
    const input=fixture(['CS'],['CS']);input.books.CS.B[1].Data.columns=input.books.CS.B[1].Data.columns.filter(column=>column!=='ResMaxAvail');
    throws(()=>publish(input),/Missing required column/,'Empty data does not conceal a broken schema');
});
scenario('optional reporting diagnostics do not become required business fields',()=>{
    const input=fixture();input.books.CS.A2[0].Data.columns=input.books.CS.A2[0].Data.columns.filter(column=>column!=='RosteredPeriodStatus'&&column!=='A2Audit');
    equal(publish(input).Report.rows.length,32,'Core C# publication remains valid without optional metadata');
});
scenario('stale AIN outputs in the CS folder fail rather than being relabelled',()=>{
    const input=fixture();input.books.CS.B[1].Data.rows[0].Role='AIN';
    throws(()=>publish(input),/Unexpected row role in CS/,'Donor output must genuinely publish CS');
});
scenario('duplicate or missing calendar matches cannot multiply or lose dated capacity',()=>{
    const duplicate=fixture();duplicate.calendar.push(structuredClone(duplicate.calendar[0]));
    throws(()=>publish(duplicate),/Duplicate Role\/Period/,'Duplicate date join key fails');
    const missing=fixture();missing.calendar=missing.calendar.filter(row=>row.Role!=='CS'||row.Period!==2);
    throws(()=>publish(missing),/Missing Role\/Period/,'Missing role/date alignment fails');
});
scenario('missing and duplicate stage cell keys fail',()=>{
    for(const column of ['Resource','Period']) {
        const input=fixture();input.books.CS.B[1].Data.rows[0][column]=null;
        throws(()=>publish(input),/Missing resource\/period/,'Required key '+column);
    }
    const duplicate=fixture();duplicate.books.CS.B[1].Data.rows.push(structuredClone(duplicate.books.CS.B[1].Data.rows[0]));
    throws(()=>publish(duplicate),/Duplicate resource\/period/,'Duplicate final cells cannot inflate capacity');
});
scenario('disabled-role calendar duplicates do not block enabled roles',()=>{
    const input=fixture();input.calendar.push({Role:'CSH',Period:1,Date:'2026-10-26'},{Role:'CSH',Period:1,Date:'2026-10-26'});
    equal(publish(input).Capacity.rows.length,8,'Calendar checks apply to enabled role labels');
});
const summary={status:'Passed independent reference examples',referenceCases:cases.length,referenceAssertions:assertions,cases,nativeEvaluationPerformed:false,productionMExecuted:false,timingMeasured:false,limitation:'Reference JavaScript does not execute production M, prove physical I/O counts or validate saved workbook data/bindings.'};
fs.writeFileSync(path.join(scratch,'reference-summary.json'),JSON.stringify(summary,null,2));
console.log(JSON.stringify(summary,null,2));
