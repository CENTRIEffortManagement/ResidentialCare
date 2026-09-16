// Run from the repository root: node scripts/test-minuteworker-period-allocation.mjs
// Independent arithmetic fixtures and source-contract checks, NOT execution of M or Excel.
import { readFile } from 'node:fs/promises';
import assert from 'node:assert/strict';

const sourceUrl = new URL('../CLIENT/DATExx-Whiddon/UNITS/Unit1/1.%20Input/Demand-MasterRoster%20Manual%20Read.xlsx_PowerQuery.m', import.meta.url);
const source = await readFile(sourceUrl, 'utf8');
let assertions = 0;
function check(condition, message) { assertions++; assert.ok(condition, message); }
function near(actual, expected, message) { check(Math.abs(actual - expected) <= 1e-8, `${message}: ${actual} vs ${expected}`); }
const sum = values => values.reduce((a, b) => a + b, 0);
const roleConfiguration = {
    RN: { dcCategory: 'RN', minuteCategory: 'RN', directCare: 1 },
    AIN: { dcCategory: 'OTHER', minuteCategory: 'OTHERS', directCare: 1 },
    CSM: { dcCategory: 'OTHER', minuteCategory: 'OTHERS', directCare: 0.5 },
};
const category = role => roleConfiguration[role].minuteCategory;
const directCare = { RN: 1, AIN: 1, CSM: 0.5 };
const targets = { RN: 1041.117, OTHERS: 5180.46666666667 - 1041.117 };
// Explicit fixture configuration, independent of production Settings Data.
const fixtureSettings = { ShiftDuration: 7.6 };
function standardHours(hours) {
    assert.ok(typeof hours === 'number' && Number.isFinite(hours) && hours > 0, 'Invalid standard FTE hours');
    return hours;
}
const fixtureHours = standardHours(fixtureSettings.ShiftDuration);
const shifts = ['AM', 'PM', 'NS'];
const weeks = [31, 32]; // Retain original, potentially non-1/2 week identifiers.

// Explicit mapping fixtures include a deliberate contradiction of the retired
// name-based guess, a many-to-one mapping, and an intentional NA exclusion.
const explicitMappings = [
    { rosterRole: ' Registered Nurse ', dcRole: 'CARE MANAGER', dcCategory: 'OTHER', directCare: 0.5 },
    { rosterRole: 'Assistant in Nursing', dcRole: 'AIN', dcCategory: 'OTHER', directCare: 1 },
    { rosterRole: 'AINC4', dcRole: 'ain', dcCategory: 'other', directCare: 1 },
    { rosterRole: 'Registered Nurse Direct', dcRole: 'RN', dcCategory: 'RN', directCare: 1 },
    { rosterRole: 'Administration', dcRole: null, dcCategory: 'NA', directCare: null },
];
const normalise = value => value == null ? null : String(value).trim().toUpperCase();
function prepareMappings(rows) {
    return rows.map(r => ({ ...r, rosterKey: normalise(r.rosterRole), roleKey: normalise(r.dcRole), dcCategory: normalise(r.dcCategory) }));
}
function validateMappings(rows) {
    const prepared = prepareMappings(rows);
    const keyCounts = new Map();
    for (const row of prepared) keyCounts.set(row.rosterKey, (keyCounts.get(row.rosterKey) ?? 0) + 1);
    const duplicateOrBlank = [...keyCounts].filter(([key, count]) => !key || count !== 1);
    const invalidCategory = prepared.filter(r => !['RN', 'OTHER', 'NA'].includes(r.dcCategory));
    const retained = prepared.filter(r => ['RN', 'OTHER'].includes(r.dcCategory));
    const invalidRole = retained.filter(r => !r.roleKey);
    const invalidDC = retained.filter(r => typeof r.directCare !== 'number' || r.directCare <= 0 || r.directCare > 1);
    const byRole = new Map();
    for (const r of retained) {
        const signature = `${r.dcCategory}|${r.directCare}`;
        if (!byRole.has(r.roleKey)) byRole.set(r.roleKey, new Set());
        byRole.get(r.roleKey).add(signature);
    }
    const conflicts = [...byRole].filter(([, signatures]) => signatures.size !== 1);
    return { prepared, retained, duplicateOrBlank, invalidCategory, invalidRole, invalidDC, conflicts };
}
function retainedMappedMasterRoles(rows, masterRoles) {
    const retainedKeys = new Set(validateMappings(rows).retained.map(r => r.rosterKey));
    return masterRoles.filter(role => retainedKeys.has(normalise(role)));
}
const validMapping = validateMappings(explicitMappings);
check(validMapping.duplicateOrBlank.length === 0 && validMapping.invalidCategory.length === 0 &&
    validMapping.invalidRole.length === 0 && validMapping.invalidDC.length === 0 &&
    validMapping.conflicts.length === 0, 'Valid explicit mapping passes');
check(validMapping.retained.find(r => r.rosterKey === 'REGISTERED NURSE').roleKey === 'CARE MANAGER',
    'Explicit mapping overrides retired Registered Nurse guess');
check(validMapping.retained.filter(r => r.roleKey === 'AIN').length === 2, 'Many roster roles map to one DC role');
check(!validMapping.retained.some(r => r.rosterKey === 'ADMINISTRATION'), 'NA mapping is excluded from analysis');
check(validateMappings([...explicitMappings, explicitMappings[0]], []).duplicateOrBlank.length === 1, 'Duplicate roster mapping fails');
check(validateMappings([...explicitMappings, { rosterRole: 'X', dcRole: 'X', dcCategory: 'INVALID', directCare: 1 }], []).invalidCategory.length === 1, 'Unknown category fails');
check(validateMappings([...explicitMappings, { rosterRole: 'X', dcRole: null, dcCategory: 'OTHER', directCare: 1 }], []).invalidRole.length === 1, 'Blank retained DC role fails');
check(validateMappings([...explicitMappings, { rosterRole: 'X', dcRole: 'X', dcCategory: 'OTHER', directCare: 0 }], []).invalidDC.length === 1, 'Invalid direct-care percentage fails');
check(validateMappings([...explicitMappings, { rosterRole: 'X', dcRole: 'X', dcCategory: 'OTHER', directCare: '0.5' }], []).invalidDC.length === 1, 'Non-numeric direct-care percentage fails');
check(validateMappings([...explicitMappings, { rosterRole: 'X', dcRole: 'AIN', dcCategory: 'RN', directCare: 1 }], []).conflicts.length === 1, 'Conflicting category for one DC role fails');
check(JSON.stringify(retainedMappedMasterRoles(explicitMappings,
    ['Registered Nurse Direct', 'Unknown Roster Role', 'Administration'])) === JSON.stringify(['Registered Nurse Direct']),
    'Unmapped and NA Master Roster roles are excluded without validation failure');
const raw = [];
for (const week of weeks) {
    for (let day = 1; day <= 7; day++) {
        const weeklyVariation = week === 31 ? 0.5 : 1.5;
        const weekdayPattern = [1, 2, 1.1, 1.2, 1, 0.8, 0.7][day - 1];
        raw.push({ role: 'RN', week, day, shift: 'AM', hours: 7.6 * weekdayPattern * weeklyVariation });
        raw.push({ role: 'RN', week, day, shift: 'PM', hours: 3.8 });
        raw.push({ role: 'AIN', week, day, shift: 'AM', hours: 15.2 });
    }
}
raw.push({ role: 'CSM', week: 31, day: 1, shift: 'AM', hours: 7.6 });

function historyGrid(rows, shiftDuration = fixtureHours) {
    const roles = [...new Set(rows.map(r => r.role))];
    const observedWeeks = [...new Set(rows.map(r => r.week))].sort((a, b) => a - b);
    return roles.flatMap(role => observedWeeks.flatMap(week => {
        const complete = new Set(rows.filter(r => r.week === week).map(r => r.day)).size === 7;
        return Array.from({ length: 7 }, (_, i) => i + 1).flatMap(day => shifts.map(shift => {
            const observed = rows.filter(r => r.role === role && r.week === week && r.day === day && r.shift === shift);
            const invalid = observed.some(r => r.hours === null || r.hours < 0);
            const status = invalid ? 'ERROR' : observed.length ? 'OBSERVED' : complete ? 'ZERO' : 'MISSING';
            const hours = invalid || status === 'MISSING' ? null : sum(observed.map(r => r.hours));
            const fortnightWeek = observedWeeks.indexOf(week) + 1;
            const fortnightDayIndex = (fortnightWeek - 1) * 7 + day;
            const fortnightDay = `${fortnightWeek}-${['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'][day - 1]}`;
            return { role, week, day, shift, fortnightWeek, fortnightDayIndex, fortnightDay, complete, status, hours, fte: hours === null ? null : hours / shiftDuration };
        }));
    }));
}
function allocate(rows, shiftDuration = fixtureHours) {
    const grid = historyGrid(rows, shiftDuration);
    assert.ok(grid.every(r => r.complete && r.hours !== null), 'Incomplete/invalid history cannot publish');
    const observedWeeks = new Set(grid.map(r => r.week)).size;
    assert.equal(observedWeeks, 2, 'Exactly two source weeks are required');
    const cells = grid.map(cell => ({ ...cell, productive: cell.hours * directCare[cell.role], category: category(cell.role) }));
    for (const cell of cells) {
        const denominator = sum(cells.filter(r => r.category === cell.category).map(r => r.productive));
        cell.share = cell.productive / denominator;
        cell.minutes = targets[cell.category] * 60 * cell.share;
        cell.fte = (cell.minutes / directCare[cell.role] / 60) / shiftDuration;
    }
    return { grid, cells };
}
const result = allocate(raw);
near(result.grid.length, 3 * 14 * 3, 'Complete unaveraged grain');
near(sum(result.grid.map(r => r.hours)), sum(raw.map(r => r.hours)), 'Historical hours preserved');
near(sum(result.grid.map(r => r.fte)) * fixtureHours, sum(raw.map(r => r.hours)), 'Historical FTE uses roster hours');
check(result.grid.some(r => r.week === 32 && r.role === 'CSM' && r.status === 'ZERO'), 'Absent role/week is explicit zero');
near(result.cells.find(r => r.role === 'CSM' && r.fortnightDay === '1-Monday' && r.shift === 'AM').hours, 7.6, 'First Monday retains its own hours');
near(result.cells.find(r => r.role === 'CSM' && r.fortnightDay === '2-Monday' && r.shift === 'AM').hours, 0, 'Second Monday remains zero, not averaged to 3.8');
near(new Set(result.cells.map(r => r.fortnightDayIndex)).size, 14, 'Fourteen independent day keys');
near(new Set(result.cells.map(r => `${r.fortnightDay}-${r.shift}`)).size, 42, 'Matrix retains 42 distinct day/shift keys');
check(result.cells.some(r => r.fortnightDay === '1-Tuesday' && r.fortnightDayIndex === 2), 'First Tuesday label and sort key');
check(result.cells.some(r => r.fortnightDay === '2-Tuesday' && r.fortnightDayIndex === 9), 'Second Tuesday label and sort key');
for (const cat of ['RN', 'OTHERS']) {
    const cells = result.cells.filter(r => r.category === cat);
    near(sum(cells.map(r => r.share)), 1, `${cat}: whole-period shares total one`);
    near(sum(cells.map(r => r.fte * fixtureHours * directCare[r.role])), targets[cat], `${cat}: original fortnight hours reconstructed without doubling`);
    for (const cell of cells) {
        const pooledCell = sum(raw.filter(r => r.role === cell.role && r.week === cell.week && r.day === cell.day && r.shift === cell.shift).map(r => r.hours * directCare[r.role]));
        const pooledCategory = sum(raw.filter(r => category(r.role) === cat).map(r => r.hours * directCare[r.role]));
        near(cell.minutes, targets[cat] * 60 * pooledCell / pooledCategory, 'Specific week/day/shift receives its share of the full fortnight');
        const dayCells = cells.filter(r => r.fortnightDayIndex === cell.fortnightDayIndex);
        const dayProductive = sum(dayCells.map(r => r.productive));
        const roleCells = dayCells.filter(r => r.role === cell.role);
        const roleProductive = sum(roleCells.map(r => r.productive));
        const roleHours = sum(roleCells.map(r => r.hours));
        if (roleHours > 0) {
            const weekdayTarget = targets[cat] * 60 * dayProductive / sum(cells.map(r => r.productive));
            near(cell.minutes, weekdayTarget * roleProductive / dayProductive * cell.hours / roleHours, 'Two-stage role/day/shift allocation reconciles');
        }
    }
}
const rnDay = (allocation, dayIndex) => sum(allocation.filter(r => r.role === 'RN' && r.fortnightDayIndex === dayIndex).map(r => r.fte));
near(sum(result.cells.filter(r => r.role === 'RN').map(r => r.fte)) / 14, targets.RN / 14 / (fixtureHours), 'BD RN mean remains an informational 14-day average');
check(rnDay(result.cells, 2) > rnDay(result.cells, 1), 'Busier Tuesday retained');
check(rnDay(result.cells, 9) > rnDay(result.cells, 2), 'Second Tuesday remains higher than first Tuesday');
const rnWeeks = weeks.map(week => sum(result.cells.filter(r => r.role === 'RN' && r.week === week).map(r => r.fte)));
check(rnWeeks[1] > rnWeeks[0], 'Week totals are allowed to differ');
near(sum(rnWeeks) * fixtureHours, targets.RN, 'Both unequal RN weeks reconcile together');
const changedTuesday = allocate(raw.map(r => r.role === 'RN' && r.week === 32 && r.day === 2 ? { ...r, hours: r.hours * 2 } : r));
check(rnDay(changedTuesday.cells, 1) < rnDay(result.cells, 1), 'Second Tuesday changes the common denominator and first Monday absolute FTE');
const incomplete = historyGrid(raw.filter(r => !(r.week === 32 && r.day === 7)));
check(incomplete.some(r => r.status === 'MISSING' && r.fte === null), 'Missing full weekday is not healthy zero');
const invalid = historyGrid([...raw, { ...raw[0], hours: -1 }, { ...raw[1], hours: null }]);
check(invalid.filter(r => r.status === 'ERROR').length === 2, 'Negative/null rows cannot be masked by aggregation');
assert.throws(() => allocate([...raw, ...raw.filter(r => r.week === 31).map(r => ({ ...r, week: 33 }))]), /Exactly two/);
assertions++;
assert.throws(() => allocate(raw.filter(r => r.week === 31)), /Exactly two/);
assertions++;
assert.throws(() => allocate(raw.filter(r => !(r.week === 32 && r.day === 7))), /Incomplete/);
assertions++;
assert.throws(() => allocate([...raw, { ...raw[0], hours: -1 }]), /Incomplete\/invalid/);
assertions++;

// Independent profile-join fixture: use actual allocations, not a generated
// scaled historical line, so misalignment and missing joins can be detected.
const allocationRows = result.cells.filter(r => r.hours > 0);
function compareProfiles(history, allocations, budgets = targets, shiftDuration = fixtureHours) {
    return history.map(h => {
        const cat = category(h.role);
        const expectedHours = budgets[cat];
        const historyHours = sum(history.filter(r => category(r.role) === cat).map(r => (r.hours ?? 0) * directCare[r.role]));
        const matches = allocations.filter(r => r.role === h.role && r.week === h.week && r.day === h.day && r.shift === h.shift);
        const actual = expectedHours == null || !h.complete || h.fte == null ? null :
            matches.length === 1 ? matches[0].fte : matches.length === 0 && h.fte === 0 ? 0 : null;
        const factor = expectedHours == null || historyHours <= 0 ? null : expectedHours / historyHours;
        const variance = h.fte === 0 && actual === 0 ? 0 : actual == null || factor == null ? null : actual - h.fte * factor;
        const status = !h.complete || h.fte == null ? 'ERROR' : expectedHours == null ? 'NO TARGET' :
            matches.length > 1 || actual == null || variance == null ? 'ERROR' :
            Math.abs(variance) <= (0.000001 / 60) / shiftDuration ? h.fte === 0 ? 'PASS ZERO' : 'PASS' : 'ERROR';
        return { ...h, actual, factor, variance, status };
    });
}
const profiles = compareProfiles(result.grid, allocationRows);
near(profiles.length, result.grid.length, 'Profile join preserves historical row count');
for (const p of profiles) {
    check(p.status.startsWith('PASS'), 'Historical and actual redistributed profile aligns');
    near(p.actual, p.fte * p.factor, 'Constant multiplier applies to every role/week/day/shift including zero cells');
}
for (const cat of ['RN', 'OTHERS']) {
    const factors = profiles.filter(p => category(p.role) === cat).map(p => p.factor);
    near(Math.max(...factors), Math.min(...factors), 'One category scalar across both weeks, all roles and shifts');
}
const changedProfile = allocationRows.map((r, i) => i === 0 ? { ...r, fte: r.fte + 0.1 } : i === 1 ? { ...r, fte: r.fte - 0.1 } : r);
check(compareProfiles(result.grid, changedProfile).filter(p => p.status === 'ERROR').length === 2, 'Shape distortion detected even when total FTE is preserved');
check(compareProfiles(result.grid, allocationRows.slice(1)).some(p => p.status === 'ERROR' && p.actual === null), 'Missing positive allocation is not silently zero-filled');
const duplicatedProfile = compareProfiles(result.grid, [...allocationRows, allocationRows[0]]);
near(duplicatedProfile.length, result.grid.length, 'Duplicate allocation does not multiply graph rows');
check(duplicatedProfile.some(p => p.status === 'ERROR'), 'Duplicate join is an error');
const noTarget = compareProfiles(result.grid, allocationRows, { RN: targets.RN });
check(noTarget.filter(p => p.role !== 'RN').every(p => p.status === 'NO TARGET' && p.actual === null), 'Absent category target remains null');
check(compareProfiles(incomplete, allocationRows).some(p => p.status === 'ERROR' && p.actual === null), 'Incomplete history is not graph-ready');
const zeroBudget = compareProfiles(result.grid, allocationRows.map(r => ({ ...r, fte: 0 })), { RN: 0, OTHERS: 0 });
check(zeroBudget.every(p => p.status.startsWith('PASS') && p.actual === 0), 'Zero target is a valid zero multiplier');
const zeroCell = result.grid.find(r => r.hours === 0);
check(compareProfiles(result.grid, [...allocationRows, { ...zeroCell, fte: 1 }]).some(p => p.status === 'ERROR'), 'Nonzero redistribution into zero-history cell fails');

// Changing the FTE unit must leave productive minutes, shares and recovered hours unchanged.
for (const hours of [7.6, 7.5, 8]) {
    const shiftDuration = standardHours(hours);
    const alternative = allocate(raw, shiftDuration);
    near(sum(alternative.grid.map(r => r.fte)) * shiftDuration, sum(raw.map(r => r.hours)), 'Historical hours round trip');
    for (let i = 0; i < alternative.cells.length; i++) {
        const cell = alternative.cells[i], baseline = result.cells[i];
        near(cell.share, baseline.share, 'Distribution independent of FTE unit');
        near(cell.minutes, baseline.minutes, 'Productive minutes independent of FTE unit');
        near(cell.fte, cell.minutes / directCare[cell.role] / (shiftDuration * 60), 'Hours-only allocation equals previous calculation');
        near(cell.fte * shiftDuration * directCare[cell.role], baseline.minutes / 60, 'Target hours round trip');
    }
    check(compareProfiles(alternative.grid, alternative.cells.filter(r => r.hours > 0), targets, shiftDuration)
        .every(p => p.status.startsWith('PASS')), 'Profiles align at each configured duration');
}
for (const invalidHours of [null, undefined, '7.6', true, 0, -1, NaN, Infinity, -Infinity]) {
    assert.throws(() => standardHours(invalidHours), /Invalid standard/);
    assertions++;
}

// Demand reads saved FTE and its conversion basis, then divides recovered hours
// by each role/shift's actual span. Exercise both repeated fortnights.
function demandRows(cells, savedHours, currentHours) {
    assert.ok(typeof savedHours === 'number' && Number.isFinite(savedHours) && savedHours > 0 &&
        Math.abs(savedHours - currentHours) <= 1e-7 / 60, 'Saved FTE basis mismatch');
    return [0, 14].flatMap(offset => cells.map(cell => {
        const span = { AM: 8.25, PM: 7.25, NS: 9 }[cell.shift];
        const hours = cell.fte * currentHours;
        return { day: cell.fortnightDayIndex + offset, hours, attendance: hours / span, span };
    }));
}
const baselineDemand = demandRows(result.cells, fixtureHours, fixtureHours);
check(baselineDemand.length === 252, '28-day demand has 252 role/day/shift cells');
for (const hours of [7.6, 7.5, 8]) {
    const shiftDuration = standardHours(hours);
    const rows = demandRows(allocate(raw, shiftDuration).cells, shiftDuration, shiftDuration);
    for (let i = 0; i < rows.length; i++) {
        near(rows[i].hours, baselineDemand[i].hours, 'Demand roster hours independent of standard FTE duration');
        near(rows[i].attendance, baselineDemand[i].attendance, 'Actual shift attendance independent of standard FTE duration');
        near(rows[i].attendance * rows[i].span, rows[i].hours, 'Demand shift integration recovers roster hours');
    }
}
for (const basis of [undefined, null, '7.6', NaN, Infinity, 0, -1, fixtureHours + 1, fixtureHours * 60]) {
    assert.throws(() => demandRows(result.cells, basis, fixtureHours), /Saved FTE basis/);
    assertions++;
}

// Source contracts tie fixtures to the selected sidecar; these are not a full M parser.
function balancedMDelimiters(text) {
    const stack = [];
    const pairs = { ')': '(', ']': '[', '}': '{' };
    let inString = false;
    let inLineComment = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        const next = text[i + 1];
        if (inLineComment) {
            if (ch === '\n') inLineComment = false;
            continue;
        }
        if (inString) {
            if (ch === '"' && next === '"') { i++; continue; }
            if (ch === '"') inString = false;
            continue;
        }
        if (ch === '/' && next === '/') { inLineComment = true; i++; continue; }
        if (ch === '"') { inString = true; continue; }
        if ('([{'.includes(ch)) stack.push(ch);
        if (')]}'.includes(ch) && stack.pop() !== pairs[ch]) return false;
    }
    return !inString && stack.length === 0;
}
const definitions = [...source.matchAll(/^shared\s+(?:#"([^"]+)"|(\w+))\s*=/gm)];
const names = definitions.map(m => m[1] || m[2]);
check(balancedMDelimiters(source), 'M source has balanced delimiters outside strings and comments');
check(new Set(names).size === names.length, 'Unique shared query names');
check(names.includes('MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE'), 'Historical output exposed');
check(/\[CategoryTargetMinutes\] \* \[#"CategoryFortnightRoleDayShiftDistribution%"\]/.test(source), 'Allocator uses full fortnight target and unaveraged shares');
check(!source.includes('[CategoryTargetMinutes] / 2'), 'Allocator does not halve the fortnight target');
check(!/\[FTE\] \* (?:7\.6|456) \* \[#"Direct Care %"\] \* 2/.test(source), 'Reconciliation does not double 14-day totals');
check(!source.includes('HistoricalWeeksInAverage'), 'No historical averaging denominator remains');
check(source.includes('Historical target facility contains exactly two source weeks'), 'Two-week validation gate present');
check(source.includes('"FortnightDayIndex"'), 'Numeric 14-day sort key present');
check(source.includes('[AllocatedDayProductiveMinutes] - [CategoryWeekdayTargetMinutes]'), 'Daily check uses weighted weekday target');
check(!source.includes('[AllocatedDayProductiveMinutes] - [CategoryDailyTargetMinutes]'), 'No fixed-category-daily reconciliation remains');
check(source.includes('each [HistoricalRosterHours] / ShiftDuration'), 'Historical FTE uses configured standard duration');
check(source.includes('Historical raw roster hours are valid'), 'Raw invalid-hour gate present');
check(source.includes('Historical graph table preserves source hours and rows'), 'Historical reconstruction gate present');
check(source.includes('Category whole-period distribution totals 100%'), 'Whole-period share gate present');
check(source.includes('shared #"MW Fortnight Hours Check"'), 'Independent raw target-hours reconciliation retained');
check(source.includes('"RedistributedRosterFTE"'), 'Paired redistributed FTE exposed');
check(source.includes('"ExpectedRedistributionFactor"') && source.includes('"ProfileVarianceFTE"'), 'Independent scalar and residual diagnostics exposed');
check(source.includes('"Facility", "MinuteCategory", "Role", "Week No", "FortnightDayIndex", "Shift"'), 'Comparison joins retain role/week/day/shift identity');
check(source.includes('"RedistributionMatchCount"'), 'Duplicate and missing allocation joins checked');
check(source.includes('Table.Combine({PreAllocationChecks, #"MW Daily Allocation Check", #"MW Fortnight Hours Check", #"MW FTE Profile Alignment Check"})'), 'Profile errors block publication');
const inputRoleQuery = source.slice(source.indexOf('// Query: INPUT MinuteWorkers'), source.indexOf('// Query: MW Roster Role Mapping Prepare'));
check(inputRoleQuery.includes('MatchingRosterRoleswithANACCRoles'), 'Replacement workbook table is authoritative');
check(!inputRoleQuery.includes('Assigned Required Types'), 'Role input returns Source directly without assigned types');
check(source.includes('shared #"MW Roster Role Mapping Prepare"'), 'Explicit roster-role mapping stage exists');
check(source.includes('"DirectCareInputIsNumeric"'), 'Mapping stage flags non-numeric direct-care inputs');
check(source.includes('then Number.From(_) else null'), 'Invalid direct-care values are made safe before calculations');
check(!source.includes('shared #"MW Abbreviate Role"'), 'Guessed abbreviation function is retired');
check(!source.includes('QFR Category'), 'Obsolete QFR Category interface is removed');
check(source.includes('each List.Contains({"RN", "OTHER"}, [DC Category])'), 'Only mapped RN/OTHER rows enter analysis');
check(!source.includes('Master Roster role has an explicit mapping'), 'Unmapped Master Roster roles do not fail validation');
check(source.includes('DC Role attributes are consistent'), 'Many-to-one DC-role attributes are validated');
const demandSource = await readFile(new URL('../CLIENT/DATExx-Whiddon/UNITS/Unit1/1.%20Input/2-DemandExtract.xlsx_PowerQuery.m', import.meta.url), 'utf8');
check(balancedMDelimiters(demandSource), 'Demand source delimiters balance');
const costSource = await readFile(new URL('../CLIENT/DATExx-Whiddon/2.%20Calculations/Cost/Cost..xlsx_PowerQuery.m', import.meta.url), 'utf8');
for (const m of [source, demandSource, costSource]) {
    const queryNames = [...m.matchAll(/^shared\s+(?:#"([^"]+)"|(\w+))\s*=/gm)].map(match => match[1] || match[2]);
    check(new Set(queryNames).size === queryNames.length, 'Unique query names in each revised source');
    check(!m.includes('StandardFTEMinutes'), 'No intermediate standard-minutes variable remains');
    check(!/\b7\.6\b|\b456\b/.test(m), 'No fixed standard FTE hours/minutes remain in edited M');
    check(m.includes('shared #"IMPORT Settings Data"'), 'Settings has a named import');
    check(m.includes('shared ShiftDuration ='), 'Standard duration is exposed in hours');
    check(m.includes('Hours <= 0') && m.includes('Number.IsNaN(Hours)') && m.includes('#infinity'), 'Invalid settings fail');
    check(m.includes('Table.RowCount(RequiredColumn) = 1'), 'Exactly one setting required');
}
check(source.includes('"Shift Net Length"'), 'Unit1 retains approved source net hours');
check(demandSource.includes('BasisMatches') && demandSource.includes('else if not BasisMatches then error'), 'Stale saved FTE basis blocks demand');
check(demandSource.includes('[SourceFTE] * ShiftDuration'), 'Demand reverses configured FTE');
check(demandSource.includes('[DemandHRS] / [DurationOfShifts]'), 'Actual shift denominator retained');
check(costSource.includes('shared ShiftHrs = ShiftDuration meta'), 'Cost keeps the hours compatibility alias');
check((costSource.match(/\*ShiftDuration\)/g) || []).length === 3, 'All three cost formulas use hours directly');
check(source.includes('([WeekdayShiftRosterMinutes] / 60) / ShiftDuration'), 'Allocation converts target minutes to hours only');
check(source.includes('(0.000001 / 60) / ShiftDuration'), 'Profile tolerance retains equivalent precision');
check(source.includes('"ShiftDuration", each ShiftDuration'), 'Manual publishes the hours basis');
check(demandSource.includes('Number.Abs(_ - ShiftDuration) <= (0.0000001 / 60)'), 'Saved basis tolerance is in hours');
for (const shiftDuration of [7.6, 7.5, 8]) {
    // Preserve costing for the same FTE and rates; equivalent to the former round trip.
    for (const [fte, rate] of [[2.25, 23], [1.125, 41.5], [3.5, 0]]) {
        near(fte * shiftDuration * rate, fte * (shiftDuration * 60 / 60) * rate, 'Hours-only cost equals previous formula');
    }
    near((0.000001 / 60) / shiftDuration, 0.000001 / (shiftDuration * 60), 'Tolerance precision preserved');
}
near(standardHours(Number.MAX_VALUE), Number.MAX_VALUE, 'Finite hours require no intermediate minutes conversion');
console.log(`PASS: ${assertions} independent arithmetic/source-contract assertions; ${names.length} shared queries. Excel runtime validation remains required.`);
