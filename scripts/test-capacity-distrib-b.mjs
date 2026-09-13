// Run from the repository root: node scripts/test-capacity-distrib-b.mjs
// JavaScript arithmetic oracle and M source-contract checks only: this does not
// execute M, read workbooks, synchronize definitions, or measure Excel refreshes.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const sourceUrl = new URL('../CLIENT/DATExx-Whiddon/UNITS/Unit1/2.%20Calculations/AIN/CapacityDistrib(B)-shifts.xlsx_PowerQuery.m', import.meta.url);
const fixtureUrl = new URL('../Workflows/ResidentialCare/Diagnostics/CapacityDistribB_RunningTotals_TEST.pq', import.meta.url);
const [source, fixtureSource] = (await Promise.all([readFile(sourceUrl, 'utf8'), readFile(fixtureUrl, 'utf8')]))
    .map(text => text.replace(/\r\n?/g, '\n'));
let assertions = 0;
let sourceChecks = 0;
function equal(actual, expected, message) { assertions++; assert.deepEqual(actual, expected, message); }
function check(condition, message) { assertions++; assert.ok(condition, message); }
function sourceCheck(condition, message) { sourceChecks++; assert.ok(condition, message); }

// Reference calculation deliberately re-sums each positional prefix, just as
// the old M expression does. Null is distinct from a numeric zero.
function prefixOracle(rows) {
    const values = rows.map(row => row.value);
    return rows.map(row => {
        const prefix = values.slice(row.start - 2, row.start - 2 + row.count);
        const present = prefix.filter(value => value !== null);
        return present.length === 0 ? null : present.reduce((total, value) => total + value, 0);
    });
}

// A separate executable model tests the intended running-state arithmetic.
// Passing this model is not evidence that the M engine executed the helper.
function runningModel(rows) {
    for (let position = 0; position < rows.length; position++) {
        const row = rows[position];
        if (!Number.isInteger(row.index) || row.index !== position + 2 ||
            !Number.isInteger(row.start) || row.start < 2 || row.start > row.index ||
            !Number.isInteger(row.count) || row.count !== row.index - row.start + 1 ||
            (position === 0 ? row.start !== row.index : row.start !== rows[position - 1].start && row.start !== row.index)) {
            throw new Error(`Invalid positional controls at row ${position + 1}`);
        }
    }
    const totals = [];
    let previousStart;
    let total = null;
    for (const row of rows) {
        if (row.start !== previousStart) total = row.value;
        else if (row.value !== null) total = total === null ? row.value : total + row.value;
        totals.push(total);
        previousStart = row.start;
    }
    return totals;
}

function fixtureRows(groups, cap = 1) {
    const rows = [];
    const totalCount = groups.reduce((count, values) => count + values.length, 0);
    for (const values of groups) {
        const start = rows.length + 2;
        for (const [offset, value] of values.entries()) {
            rows.push({
                rowId: `row-${String(totalCount - rows.length).padStart(3, '0')}`,
                priority: 1,
                value,
                index: rows.length + 2,
                start,
                count: offset + 1,
                cap,
                incomingKeep: rows.length % 3 === 2 ? 'Remove' : 'Keep',
            });
        }
    }
    return rows;
}

// In M, comparisons involving null return null and cannot be used directly as
// an if condition. Represent that existing decision error explicitly here.
function decisionOutcome(total, cap, incomingKeep) {
    const nullComparisonError = 'M null condition error';
    const keep = total === null ? nullComparisonError : total <= cap ? 'Keep' : 'Remove';
    return [
        total === null ? nullComparisonError : total > cap ? 'Remove' : total <= cap ? 'Keep' : 'Split',
        keep,
        incomingKeep !== 'Keep' ? null : keep === nullComparisonError ? nullComparisonError : keep === 'Keep' ? 'Keep' : null,
    ];
}

const cases = [
    { name: 'empty', groups: [], expected: [] },
    { name: 'singleton', groups: [[5]], expected: [5] },
    { name: 'singleton-null', groups: [[null]], expected: [null] },
    { name: 'multiple-groups', groups: [[1, 2, 3], [4], [2, 0, 5]], expected: [1, 3, 6, 4, 2, 2, 7] },
    { name: 'all-null-groups', groups: [[null, null], [null]], expected: [null, null, null] },
    { name: 'null-leading-interior-trailing', groups: [[null, 2, null, 0, null], [null, -1, null]], expected: [null, 2, 2, 2, 2, null, -1, -1] },
    { name: 'zero-is-not-null', groups: [[0, null, 0], [null, 0]], expected: [0, 0, 0, null, 0] },
    { name: 'negative-values', groups: [[3, -5, 2, -1], [-2, 1]], expected: [3, -2, 0, -1, -2, -1] },
    { name: 'binary-fractions', groups: [[0.25, 0.5, 0.125], [1.5, -0.25]], expected: [0.25, 0.75, 0.875, 1.5, 1.25] },
    { name: 'exact-cap-and-next-row', groups: [[0.25, 0.75, 0, 0.125]], expected: [0.25, 1, 1, 1.125] },
];
for (const test of cases) {
    const rows = fixtureRows(test.groups);
    equal(prefixOracle(rows), test.expected, `${test.name}: independent explicit oracle`);
    equal(runningModel(rows), test.expected, `${test.name}: running-state arithmetic`);
}
equal(decisionOutcome(null, 1, 'Keep'), ['M null condition error', 'M null condition error', 'M null condition error'], 'Null totals must not become zero/Keep');
equal(decisionOutcome(null, 1, 'Remove'), ['M null condition error', 'M null condition error', null], 'Incoming Remove retains the existing short circuit');
equal(decisionOutcome(1, 1, 'Keep'), ['Keep', 'Keep', 'Keep'], 'Exact cap stays Keep');
equal(decisionOutcome(1.125, 1, 'Keep'), ['Remove', 'Remove', null], 'Above cap is removed');

const tiedRows = fixtureRows([[2, 1, 4], [3, 2]], 3);
equal(runningModel(tiedRows), [2, 3, 7, 3, 5], 'Tied priorities retain supplied index order');
equal(tiedRows.map(row => row.rowId), ['row-005', 'row-004', 'row-003', 'row-002', 'row-001'], 'Fixture labels expose accidental alphabetical sorting');

const base = fixtureRows([[1, 2, 3]], 3);
const invalidRows = [
    base.map(row => ({ ...row, index: row.index - 1 })),
    base.map(row => ({ ...row, index: row.index === 3 ? 2 : row.index })),
    base.map(row => ({ ...row, index: row.index > 2 ? row.index + 1 : row.index })),
    [...base].reverse(),
    base.map(row => ({ ...row, index: row.index === 3 ? null : row.index })),
    base.map(row => ({ ...row, index: row.index === 3 ? '3' : row.index })),
    base.map(row => ({ ...row, start: 1, count: row.count + 1 })),
    fixtureRows([[1], [2], [3]]).map(row => row.index === 4 ? { ...row, start: 2, count: 3 } : row),
    base.map(row => ({ ...row, start: null })),
    base.map(row => ({ ...row, start: 5 })),
    base.map(row => ({ ...row, count: row.count + 1 })),
    base.map(row => ({ ...row, count: row.count + 0.5 })),
    base.map(row => ({ ...row, count: null })),
];
for (const [index, rows] of invalidRows.entries()) {
    assertions++;
    assert.throws(() => runningModel(rows), /Invalid positional controls/, `Invalid fixture ${index + 1} must fail`);
}

// Reproducible property checks use binary fractions to avoid treating a JS
// floating-point convention as proof of Power Query's List.Sum implementation.
let seed = 0x5a17;
function random(maximum) {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed % maximum;
}
const randomGroupSets = 300;
for (let test = 0; test < randomGroupSets; test++) {
    const groups = Array.from({ length: 1 + random(8) }, () =>
        Array.from({ length: 1 + random(60) }, () => random(7) === 0 ? null : (random(25) - 12) / 8));
    const rows = fixtureRows(groups, (random(25) - 8) / 8);
    const originalRows = structuredClone(rows);
    const expected = prefixOracle(rows);
    const actual = runningModel(rows);
    equal(actual, expected, `Generated group set ${test}: exact binary-fraction totals`);
    equal(rows, originalRows, `Generated group set ${test}: all original rows and order retained`);
    equal(actual.map((total, i) => decisionOutcome(total, rows[i].cap, rows[i].incomingKeep)),
        expected.map((total, i) => decisionOutcome(total, rows[i].cap, rows[i].incomingKeep)),
        `Generated group set ${test}: cap decisions retained`);
}

const twentyDayRows = fixtureRows(Array.from({ length: 30 }, () =>
    Array.from({ length: 20 * 3 }, (_, i) => i % 7 === 0 ? null : i % 5 === 0 ? 0 : 0.25)), 8);
equal(runningModel(twentyDayRows), prefixOracle(twentyDayRows), '20-day, 3-shift, 30-resource synthetic arithmetic');
const operationCounts = [10, 20].map(days => ({
    days,
    rows: 30 * days * 3,
    legacySummedValues: 30 * (days * 3) * (days * 3 + 1) / 2,
    runningValues: 30 * days * 3,
}));
check(operationCounts[1].legacySummedValues > 3.9 * operationCounts[0].legacySummedValues, 'Resource-prefix work increases approximately quadratically with days');
equal(operationCounts[1].runningValues, 2 * operationCounts[0].runningValues, 'One-pass arithmetic grows with row count');

// Extract a shared definition through its unquoted, uncommented terminator.
// This is a source-contract reader; a separate approved M parser is still required.
function queryText(name) {
    const declarations = [...source.matchAll(/^shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][\w]*)\s*=/gm)];
    const declaration = declarations.find(match =>
        (match[1].startsWith('#"') ? match[1].slice(2, -1).replaceAll('""', '"') : match[1]) === name);
    assert.ok(declaration, `Missing shared query ${name}`);
    let inString = false;
    let lineComment = false;
    let blockComment = false;
    for (let i = declaration.index; i < source.length; i++) {
        const current = source[i];
        const next = source[i + 1];
        if (lineComment) { if (current === '\n') lineComment = false; continue; }
        if (blockComment) { if (current === '*' && next === '/') { blockComment = false; i++; } continue; }
        if (inString) {
            if (current === '"' && next === '"') i++;
            else if (current === '"') inString = false;
            continue;
        }
        if (current === '/' && next === '/') { lineComment = true; i++; continue; }
        if (current === '/' && next === '*') { blockComment = true; i++; continue; }
        if (current === '"') { inString = true; continue; }
        if (current === ';') return source.slice(declaration.index, i + 1);
    }
    assert.fail(`Missing terminator for ${name}`);
}
const compact = text => text.replace(/\s+/g, ' ');
const helper = queryText('fnAddIndexedRunningTotal');
sourceCheck(helper.includes('List.Generate(') && helper.includes('List.Buffer('), 'Helper uses the planned iterative, buffered-list calculation');
sourceCheck(!/Table\.(?:Group|Sort)\s*\(/.test(helper), 'Helper does not reorder or regroup supplied priority rows');
sourceCheck(/Table\.AddColumn\([\s\S]*type any\s*\)/.test(helper), 'Added total retains the legacy any column type');
const mappings = [
    ['ReDistributeResAvailability', '#"Changed Type", "C#", "IndexAllRowPrioritySort", "StartResIndex", "Subtraction", "RunningResTotal"', 'IndexAllRowPrioritySort'],
    ['ReDistribPeriodAvailbility TABLE', '#"Inserted Subtraction", "ApplicableC#", "Index", "StartPeriodIndex", "Subtraction", "PeriodRunningTotal"', 'Index'],
    ['SubtractOverallatedPeriods', '#"Changed Type", "MinimumAvailable", "IndexAllRows", "StartPeriodIndex", "Subtraction", "RunningPeriodTotal"', 'IndexAllRows'],
    ['SubtractOverallocatedResources', '#"Inserted Subtraction", "MinimumAvailable", "IndexAllRows", "StartResIndex", "Subtraction", "RunningResTotal"', 'IndexAllRows'],
];
for (const [name, args, indexColumn] of mappings) {
    const query = queryText(name);
    sourceCheck(compact(query).includes(`fnAddIndexedRunningTotal(${args})`), `${name}: exact value/index/start/count/output mapping`);
    sourceCheck(!/List\.Range\s*\(/.test(query), `${name}: repeated prefix scan removed`);
    sourceCheck(query.includes(`{{"${indexColumn}", Order.Ascending}}`), `${name}: post-join positional sort retained`);
}
sourceCheck((source.match(/\bfnAddIndexedRunningTotal\s*\(/g) ?? []).length === 4, 'Exactly four production calls use the helper');

const expectedProtectedQuery = `shared #"ResPeriodAvailabilityCapped(C#)TABLE" = let
    Source = #"IMPORT ResPeriodAvailabilityCapped(C#)",
    BUFFER = Table.Buffer(Source),
    #"Renamed Columns" = Table.RenameColumns(BUFFER,{{"AvailabilityCapped", "C#"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each ([#"C#"] <> null)),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows",0,null,Replacer.ReplaceValue,{"C#"}),
    #"Filtered Rows1" = Table.SelectRows(#"Replaced Value", each ([Role] = Role) )
in
    #"Filtered Rows1";`;
sourceCheck(queryText('ResPeriodAvailabilityCapped(C#)TABLE') === expectedProtectedQuery,
    'Protected early buffer, null filter, zero-to-null conversion and Role filter remain exactly unchanged');

// These import bodies are the reviewed baseline after replacing only Source.
// Exact comparison protects navigation names, type conversions and filter order.
const importBodies = [
    ['IMPORT Demand', `    Demand_Prepare_Table = Source{[Item="Demand_Prepare",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Demand_Prepare_Table,{{"Shift", type text}, {"Role", type text}, {"Facility", type text}, {"D", Int64.Type}, {"Date", type date}, {"Period", Int64.Type}})
in
    #"Changed Type";`],
    ['IMPORT Allocation', `    ResPeriodAllocationTABLE_Table = Source{[Item="ResPeriodAllocationTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAllocationTABLE_Table,{{"Role", type text}, {"Allocation", type number}, {"Resource", Int64.Type}, {"Period", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Period] <> null))
in
    #"Filtered Rows";`],
    ['IMPORT AvailabilityOriginal', `    ResPeriodAvailabilityTABLE_Table = Source{[Item="ResPeriodAvailabilityTABLE",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(ResPeriodAvailabilityTABLE_Table,{{"Role", type text}, {"Resource", Int64.Type}, {"Period", Int64.Type}, {"Availability", Int64.Type}}),
    BUFFER = Table.Buffer(#"Changed Type")
in
    BUFFER;`],
    ['IMPORT ResPeriodNWDTABLE', `    ResPeriodWDTABLE_Sheet = Source{[Item="ResPeriodWDTABLE",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(ResPeriodWDTABLE_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Resource", Int64.Type}, {"Day", Int64.Type}, {"PotentialAvailability", type text}, {"Period", Int64.Type}, {"Availability", Int64.Type}, {"RosteredPeriodStatus", type text}})
in
    #"Changed Type";`],
];
for (const [name, body] of importBodies) {
    const expected = `shared #"${name}" = let\n    Source = #"IMPORTSource A1",\n${body}`;
    sourceCheck(queryText(name) === expected, `${name}: only the Source reference changed from the reviewed import`);
}
const a1Source = queryText('IMPORTSource A1');
sourceCheck(a1Source.includes(String.raw`Binary.Buffer(File.Contents(RolePath&"\CapacityDistrib(A.1)-shifts.xlsx"))`) &&
    a1Source.includes('Table.Buffer(Excel.Workbook(WorkbookBinary, null, true))'),
    'Shared A.1 source retains the exact saved-file target and workbook navigation options');

const prioritySetup = queryText('PrioritiseReductionAvail-Setup');
sourceCheck(compact(prioritySetup).includes('Table.NestedJoin(#"Expanded PeriodOverallocatedPriroristised", {"Resource", "Period"}, #"IMPORT ResPeriodNWDTABLE", {"Resource", "Period"}, "IMPORT ResPeriodNWDTABLE", JoinKind.LeftOuter)'),
    'Reduction priority join matches Resource to Resource and Period to Period');
sourceCheck(prioritySetup.includes('{{"Period", Order.Ascending}, {"Excess%", Order.Descending}, {"ResPeriodC##MAXReduction", Order.Descending}}'),
    'Correcting the priority join does not add a new business tie-breaker');

const finalOutput = queryText('C###TABLE B');
sourceCheck(finalOutput.includes('InputChecks = CapacityDistribB_INPUT_CHECK') &&
    compact(finalOutput).includes('Source = if List.AllTrue(InputChecks[Passed]) then #"C##TABLE" else error Error.Record(') &&
    finalOutput.includes('Table.NestedJoin(Source, {"Resource", "Period"}, SubtractOverallocatedResources'),
    'Required input validation gates the source consumed by final B capacity');
for (const name of ['C###SUM', 'C###MATRIX']) {
    sourceCheck(queryText(name).includes('Source = #"C###TABLE B"'), `${name}: consumes the gated final table`);
}
const fixtureNames = [...fixtureSource.matchAll(/\[Name = "([^"]+)"/g)].map(match => match[1]);
sourceCheck(fixtureNames.length >= 34 && new Set(fixtureNames).size === fixtureNames.length, 'M fixture coverage remains present with unique names');
sourceCheck(fixtureSource.includes('shared CapacityDistribB_RunningTotals_CHECK') && fixtureSource.includes('error Error.Record("CapacityDistribB.TestFailure"'),
    'M fixtures expose a fail-closed check query');

console.log(JSON.stringify({
    status: 'PASS',
    scope: 'JavaScript arithmetic oracle and source contracts; not M execution or an Excel refresh benchmark',
    assertions,
    sourceChecks,
    randomGroupSets,
    pureMFixtureCases: fixtureNames.length,
    syntheticOperationCounts: operationCounts,
}, null, 2));
