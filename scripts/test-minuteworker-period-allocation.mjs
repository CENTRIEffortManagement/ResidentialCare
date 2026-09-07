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
const category = role => role === 'RN' ? 'RN' : 'OTHERS';
const directCare = { RN: 1, AIN: 1, CSM: 0.5 };
const targets = { RN: 1041.117, OTHERS: 5180.46666666667 - 1041.117 };
const shifts = ['AM', 'PM', 'NS'];
const weeks = [31, 32]; // Retain original, potentially non-1/2 week identifiers.
const raw = [];
for (const week of weeks) {
    for (let day = 1; day <= 7; day++) {
        const weeklyVariation = week === 31 ? 0.9 : 1.1;
        const weekdayPattern = [1, 2, 1.1, 1.2, 1, 0.8, 0.7][day - 1];
        raw.push({ role: 'RN', week, day, shift: 'AM', hours: 7.6 * weekdayPattern * weeklyVariation });
        raw.push({ role: 'RN', week, day, shift: 'PM', hours: 3.8 });
        raw.push({ role: 'AIN', week, day, shift: 'AM', hours: 15.2 });
    }
}
raw.push({ role: 'CSM', week: 31, day: 1, shift: 'AM', hours: 7.6 });

function historyGrid(rows) {
    const roles = [...new Set(rows.map(r => r.role))];
    const observedWeeks = [...new Set(rows.map(r => r.week))];
    return roles.flatMap(role => observedWeeks.flatMap(week => {
        const complete = new Set(rows.filter(r => r.week === week).map(r => r.day)).size === 7;
        return Array.from({ length: 7 }, (_, i) => i + 1).flatMap(day => shifts.map(shift => {
            const observed = rows.filter(r => r.role === role && r.week === week && r.day === day && r.shift === shift);
            const invalid = observed.some(r => r.hours === null || r.hours < 0);
            const status = invalid ? 'ERROR' : observed.length ? 'OBSERVED' : complete ? 'ZERO' : 'MISSING';
            const hours = invalid || status === 'MISSING' ? null : sum(observed.map(r => r.hours));
            return { role, week, day, shift, complete, status, hours, fte: hours === null ? null : hours / 7.6 };
        }));
    }));
}
function allocate(rows) {
    const grid = historyGrid(rows);
    assert.ok(grid.every(r => r.complete && r.hours !== null), 'Incomplete/invalid history cannot publish');
    const observedWeeks = new Set(grid.map(r => r.week)).size;
    const averaged = grid.filter(r => r.week === grid[0].week).map(cell => {
        const hours = sum(grid.filter(r => r.role === cell.role && r.day === cell.day && r.shift === cell.shift).map(r => r.hours)) / observedWeeks;
        return { ...cell, hours, productive: hours * directCare[cell.role], category: category(cell.role) };
    });
    for (const cell of averaged) {
        const denominator = sum(averaged.filter(r => r.category === cell.category).map(r => r.productive));
        cell.share = cell.productive / denominator;
        cell.minutes = targets[cell.category] * 60 / 2 * cell.share;
        cell.fte = cell.minutes / directCare[cell.role] / 456;
    }
    return { grid, averaged };
}
const result = allocate(raw);
near(result.grid.length, 3 * 14 * 3, 'Complete unaveraged grain');
near(sum(result.grid.map(r => r.hours)), sum(raw.map(r => r.hours)), 'Historical hours preserved');
near(sum(result.grid.map(r => r.fte)) * 7.6, sum(raw.map(r => r.hours)), 'Historical FTE uses roster hours');
check(result.grid.some(r => r.week === 32 && r.role === 'CSM' && r.status === 'ZERO'), 'Absent role/week is explicit zero');
near(result.averaged.find(r => r.role === 'CSM' && r.day === 1 && r.shift === 'AM').hours, 3.8, 'Both Mondays in averaging denominator');
for (const cat of ['RN', 'OTHERS']) {
    const cells = result.averaged.filter(r => r.category === cat);
    near(sum(cells.map(r => r.share)), 1, `${cat}: whole-period shares total one`);
    near(sum(cells.map(r => r.fte * 7.6 * directCare[r.role] * 2)), targets[cat], `${cat}: original fortnight hours reconstructed`);
    for (const cell of cells) {
        const pooledCell = sum(raw.filter(r => r.role === cell.role && r.day === cell.day && r.shift === cell.shift).map(r => r.hours * directCare[r.role]));
        const pooledCategory = sum(raw.filter(r => category(r.role) === cat).map(r => r.hours * directCare[r.role]));
        near(cell.minutes, targets[cat] * 60 * pooledCell / pooledCategory / 2, 'Pooled-fortnight allocation then weekday averaging is equivalent');
        const dayCells = cells.filter(r => r.day === cell.day);
        const dayProductive = sum(dayCells.map(r => r.productive));
        const roleCells = dayCells.filter(r => r.role === cell.role);
        const roleProductive = sum(roleCells.map(r => r.productive));
        const roleHours = sum(roleCells.map(r => r.hours));
        if (roleHours > 0) {
            const weekdayTarget = targets[cat] * 60 / 2 * dayProductive / sum(cells.map(r => r.productive));
            near(cell.minutes, weekdayTarget * roleProductive / dayProductive * cell.hours / roleHours, 'Two-stage role/day/shift allocation reconciles');
        }
    }
}
const rnDay = (allocation, day) => sum(allocation.filter(r => r.role === 'RN' && r.day === day).map(r => r.fte));
near(sum(result.averaged.filter(r => r.role === 'RN').map(r => r.fte)) / 7, targets.RN / 14 / 7.6, 'BD RN average daily FTE');
check(rnDay(result.averaged, 2) > rnDay(result.averaged, 1), 'Busier Tuesday retained');
const changedTuesday = allocate(raw.map(r => r.role === 'RN' && r.day === 2 ? { ...r, hours: r.hours * 2 } : r));
check(rnDay(changedTuesday.averaged, 1) < rnDay(result.averaged, 1), 'Tuesday changes the common denominator and Monday absolute FTE');
const incomplete = historyGrid(raw.filter(r => !(r.week === 32 && r.day === 7)));
check(incomplete.some(r => r.status === 'MISSING' && r.fte === null), 'Missing full weekday is not healthy zero');
const invalid = historyGrid([...raw, { ...raw[0], hours: -1 }, { ...raw[1], hours: null }]);
check(invalid.filter(r => r.status === 'ERROR').length === 2, 'Negative/null rows cannot be masked by aggregation');
const thirdWeek = allocate([...raw, ...raw.filter(r => r.week === 31).map(r => ({ ...r, week: 33 }))]);
near(new Set(thirdWeek.grid.map(r => r.week)).size, 3, 'All supplied weeks retained');
near(sum(thirdWeek.averaged.filter(r => r.role === 'RN').map(r => r.fte)) * 7.6 * 2, targets.RN, 'Longer history still allocates a fortnight budget');

// Source contracts tie fixtures to the selected sidecar; these are not a full M parser.
const definitions = [...source.matchAll(/^shared\s+(?:#"([^"]+)"|(\w+))\s*=/gm)];
const names = definitions.map(m => m[1] || m[2]);
check(new Set(names).size === names.length, 'Unique shared query names');
check(names.includes('MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE'), 'Historical output exposed');
check(/\[CategoryTargetMinutes\] \/ 2 \*\s*\[#"CategoryWeekRoleDayShiftDistribution%"\]/.test(source), 'Allocator uses whole-period share and one divide-by-two');
check(source.includes('[AllocatedDayProductiveMinutes] - [CategoryWeekdayTargetMinutes]'), 'Daily check uses weighted weekday target');
check(!source.includes('[AllocatedDayProductiveMinutes] - [CategoryDailyTargetMinutes]'), 'No fixed-category-daily reconciliation remains');
check(source.includes('each [HistoricalRosterHours] / 7.6'), 'Historical FTE uses 7.6-hour denominator');
check(source.includes('Historical raw roster hours are valid'), 'Raw invalid-hour gate present');
check(source.includes('Historical graph table preserves source hours and rows'), 'Historical reconstruction gate present');
check(source.includes('Category whole-period distribution totals 100%'), 'Whole-period share gate present');
check(source.includes('shared #"MW Fortnight Hours Check"'), 'Independent raw target-hours reconciliation retained');
console.log(`PASS: ${assertions} independent arithmetic/source-contract assertions; ${names.length} shared queries. Excel runtime validation remains required.`);
