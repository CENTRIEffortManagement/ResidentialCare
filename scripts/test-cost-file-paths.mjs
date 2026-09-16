// Independent path fixtures and source contracts; does not inspect Excel or execute M.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const source = await readFile(new URL('../CLIENT/DATExx-Whiddon/2.%20Calculations/Cost/Cost..xlsx_PowerQuery.m', import.meta.url), 'utf8');
let checks = 0;
const normalize = value => value == null ? null : value.replaceAll('/', '\\').replace(/\\+$/, '');
const fixtureRoot = String.raw`D:\Fixture\ResidentialCare\CLIENT\DATExx-Whiddon`;
const suffix = String.raw`\2. Calculations\Cost\Cost..xlsx`;
const mapping = [{ SharepointRootUrl: 'https://example.test/sites/Care', SyncedFolderRootPath: 'D:/Fixture' }];

function clientRoot(rows, mappings) {
    assert.equal(rows.length, 1, 'one FilePathUrl row');
    const row = rows[0];
    let path = Object.hasOwn(row, 'FilePath') ? row.FilePath :
        Object.keys(row).length === 1 && Object.hasOwn(row, 'Column1') ? row.Column1 : null;
    assert.ok(typeof path === 'string' && path.trim(), 'nonblank FilePathUrl');
    path = path.trim().replaceAll('/', '\\');
    if (path.includes('[')) {
        assert.ok(path.includes(']'), 'CELL filename brackets');
        path = path.slice(0, path.indexOf('[')) + path.slice(path.indexOf('[') + 1, path.indexOf(']'));
    }
    assert.equal(path.slice(path.lastIndexOf('\\') + 1).toLowerCase(), 'cost..xlsx', 'workbook identity');
    path = normalize(path);
    const candidates = mappings.flatMap(m => {
        const local = normalize(m.SyncedFolderRootPath), url = normalize(m.SharepointRootUrl);
        return [url, url == null ? null : url + '\\Shared Documents', local]
            .filter(root => root && root.trim() && local && local.trim())
            .map(root => ({ root, local }));
    }).filter(m => path.toLowerCase().startsWith(m.root.toLowerCase()) &&
        (path.length === m.root.length || path[m.root.length] === '\\'))
        .sort((a, b) => b.root.length - a.root.length);
    assert.ok(candidates.length, 'matching root');
    const match = candidates[0];
    const localPath = match.local + path.slice(match.root.length);
    const folder = localPath.slice(0, localPath.lastIndexOf('\\'));
    const tail = '\\2. Calculations\\Cost';
    assert.ok(folder.toLowerCase().endsWith(tail.toLowerCase()), 'Cost folder');
    const root = folder.slice(0, -tail.length);
    assert.ok(root.toLowerCase().endsWith('\\datexx-whiddon'), 'Whiddon client');
    return root;
}
function expectPath(value, maps, expected) {
    assert.equal(clientRoot([value], maps), expected);
    checks++;
}
expectPath({ FilePath: fixtureRoot + suffix }, mapping, fixtureRoot);
expectPath({ FilePath: (fixtureRoot + suffix).replaceAll('\\', '/') }, mapping, fixtureRoot);
expectPath({ Column1: fixtureRoot + '\\2. Calculations\\Cost\\[Cost..xlsx]Sheet1' }, mapping, fixtureRoot);
expectPath({ FilePath: 'https://example.test/sites/Care/ResidentialCare/CLIENT/DATExx-Whiddon/2. Calculations/Cost/Cost..xlsx' }, mapping, fixtureRoot);
expectPath({ FilePath: 'https://example.test/sites/Care/Shared Documents/ResidentialCare/CLIENT/DATExx-Whiddon/2. Calculations/Cost/[Cost..xlsx]Sheet1' }, mapping, fixtureRoot);
expectPath({ FilePath: fixtureRoot + suffix }, [{ SharepointRootUrl: null, SyncedFolderRootPath: 'D:/Fixture/' }], fixtureRoot);
expectPath({ FilePath: fixtureRoot + suffix }, [
    ...mapping,
    { SharepointRootUrl: 'D:/Fixture/ResidentialCare', SyncedFolderRootPath: 'E:/Mapped/ResidentialCare' }
], 'E:\\Mapped\\ResidentialCare\\CLIENT\\DATExx-Whiddon');
expectPath({ FilePath: (fixtureRoot + suffix).toLowerCase() }, mapping, 'D:\\Fixture\\residentialcare\\client\\datexx-whiddon');
expectPath({ FilePath: '\\\\server\\care\\CLIENT\\DATExx-Whiddon' + suffix },
    [{ SharepointRootUrl: null, SyncedFolderRootPath: '\\\\server\\care' }],
    '\\\\server\\care\\CLIENT\\DATExx-Whiddon');

for (const rows of [
    [], [{ FilePath: '' }], [{ FilePath: null }], [{ WrongColumn: fixtureRoot + suffix }],
    [{ FilePath: fixtureRoot + suffix }, { FilePath: fixtureRoot + suffix }],
    [{ FilePath: (fixtureRoot + suffix).replace('Cost..xlsx', 'Copy-Cost..xlsx') }],
    [{ FilePath: (fixtureRoot + suffix).replace('DATExx-Whiddon', 'DATExx') }],
    [{ FilePath: (fixtureRoot + suffix).replace('\\Cost\\', '\\Other\\') }],
    [{ FilePath: (fixtureRoot + suffix).replace('Fixture', 'FixtureOther') }],
    [{ FilePath: (fixtureRoot + suffix).replace('D:', 'Z:') }],
]) {
    assert.throws(() => clientRoot(rows, mapping));
    checks++;
}

const contracts = [
    !source.includes('[Name="Folder"]'),
    (source.match(/\[Name="FilePathUrl"\]/g) || []).length === 1,
    source.includes('shared #"IMPORT CentriSyncPaths"'),
    source.includes('shared CostPathTABLE ='),
    source.includes('shared ClientPath ='),
    source.includes('Text.StartsWith(FilePath, [MatchRoot], Comparer.OrdinalIgnoreCase)'),
    source.includes('Text.Range(FilePath, [MatchRootLength], 1) = "\\"'),
    source.includes('Table.Sort(MatchingRows, {{"MatchRootLength", Order.Descending}})'),
    source.includes('ClientPath & "\\UNITS\\Unit1\\2. Calculations\\Settings Data.xlsx"'),
    source.includes('ClientPath & "\\2. Calculations\\E-O-I\\Inefficiencies.xlsx"'),
    source.includes('Source = #"IMPORT Inefficiencies"'),
    (source.match(/File\.Contents\(/g) || []).length === 3,
    !source.includes('FACILITIES\\ASHB'),
    !source.includes("Cliff's Computer"),
];
for (const valid of contracts) { assert.ok(valid, 'Cost path source contract'); checks++; }
console.log('PASS: ' + checks + ' Cost path fixture/source checks. No Excel or M runtime execution.');
