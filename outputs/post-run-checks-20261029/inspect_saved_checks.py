"""Read saved worksheet diagnostics and cell errors; never open Excel or write XLSX."""
import collections
import hashlib
import json
import posixpath
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
DATE = ROOT / 'CLIENT/DATExx-Whiddon'
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
S = '{' + NS['s'] + '}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
CHECK = re.compile(r'check|validat|reconcil|diagnos|issue|audit|test|error|exception|mismatch|missing|duplicate|integrity|quality', re.I)

def digest(path):
    with path.open('rb') as f:
        return hashlib.file_digest(f, 'sha256').hexdigest()

def rels(z, part):
    folder, name = posixpath.split(part)
    key = posixpath.join(folder, '_rels', name + '.rels')
    if key not in z.namelist():
        return {}
    return {r.attrib['Id']: (r.attrib['Target'].lstrip('/') if r.attrib['Target'].startswith('/')
            else posixpath.normpath(posixpath.join(folder, r.attrib['Target'])))
            for r in ET.fromstring(z.read(key)) if r.attrib.get('TargetMode') != 'External'}

def colnum(text):
    n = 0
    for ch in text:
        n = n * 26 + ord(ch) - 64
    return n

def bounds(ref):
    cells = ref.split(':')
    a, b = cells[0], cells[-1]
    x = re.fullmatch(r'([A-Z]+)(\d+)', a)
    y = re.fullmatch(r'([A-Z]+)(\d+)', b)
    return colnum(x[1]), int(x[2]), colnum(y[1]), int(y[2])

def value(c, shared):
    kind = c.attrib.get('t')
    v = c.find(S + 'v')
    if kind == 'inlineStr':
        return ''.join(c.find(S + 'is').itertext())
    if v is None or v.text is None:
        return None
    if kind == 's':
        return shared[int(v.text)]
    if kind == 'b':
        return v.text == '1'
    if kind in ('e', 'str', 'd'):
        return v.text
    try:
        return float(v.text) if any(ch in v.text for ch in '.eE') else int(v.text)
    except ValueError:
        return v.text

def inspect(job):
    path = (DATE / job['Path']).resolve()
    before = digest(path)
    result = {'id': job['Id'], 'path': job['Path'], 'run': job['Run'],
              'metadata': job['Metadata']['OutputStatus'], 'sha256': before,
              'tables': [], 'diagnostics': [], 'cell_errors': [],
              'broken_defined_names': [], 'unloaded_check_connections': []}
    with ZipFile(path, 'r') as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            shared = [''.join(si.itertext()) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        for defined in wb.findall('s:definedNames/s:definedName', NS):
            if '#REF!' in (defined.text or ''):
                result['broken_defined_names'].append({'name': defined.attrib['name'], 'value': defined.text})
        connections = {}
        if 'xl/connections.xml' in z.namelist():
            for conn in ET.fromstring(z.read('xl/connections.xml')):
                connections[conn.attrib['id']] = conn.attrib
        loaded = set()
        for part in z.namelist():
            if re.fullmatch(r'xl/queryTables/queryTable\d+\.xml', part):
                loaded.add(ET.fromstring(z.read(part)).attrib.get('connectionId'))
        result['unloaded_check_connections'] = [c['name'] for cid, c in connections.items()
                if CHECK.search(c.get('name', '')) and cid not in loaded and c.get('deleted') not in ('1', 'true')]
        bookrels = rels(z, 'xl/workbook.xml')
        for sheet in wb.find(S + 'sheets'):
            name = sheet.attrib['name']
            sheetpath = bookrels[sheet.attrib[REL + 'id']]
            sheetrels = rels(z, sheetpath)
            tables = []
            for target in sheetrels.values():
                if '/tables/' not in target:
                    continue
                table = ET.fromstring(z.read(target))
                cols = [c.attrib['name'] for c in table.find(S + 'tableColumns')]
                left, top, right, bottom = bounds(table.attrib['ref'])
                info = {'name': table.attrib.get('displayName', table.attrib.get('name')),
                        'sheet': name, 'ref': table.attrib['ref'], 'columns': cols,
                        'declared_rows': max(0, bottom-top-int(table.attrib.get('totalsRowCount', '0'))),
                        'observed_rows': 0}
                diagnostic = bool(CHECK.search(info['name']) or CHECK.search(name) or
                                  ('Status' in cols and any(c in cols for c in ('Check', 'Test', 'Rule'))))
                info['diagnostic'] = diagnostic
                result['tables'].append(info)
                tables.append((left, top, right, bottom-int(table.attrib.get('totalsRowCount', '0')), info))
                if diagnostic:
                    info['data'] = []
                    result['diagnostics'].append(info)
            with z.open(sheetpath) as stream:
                for _, row in ET.iterparse(stream, events=('end',)):
                    if row.tag != S + 'row':
                        continue
                    index = int(row.attrib['r'])
                    active = [t for t in tables if t[1] < index <= t[3]]
                    cells = {}
                    for cell in row.findall(S + 'c'):
                        coord = cell.attrib['r']
                        if cell.attrib.get('t') == 'e':
                            result['cell_errors'].append({'sheet': name, 'cell': coord, 'error': value(cell, shared)})
                        if active:
                            cells[colnum(re.match(r'[A-Z]+', coord)[0])] = cell
                    for left, top, right, bottom, info in active:
                        vals = {c: value(cells[left+i], shared) if left+i in cells else None
                                for i, c in enumerate(info['columns'])}
                        if any(v is not None for v in vals.values()):
                            info['observed_rows'] += 1
                            if info['diagnostic']:
                                info['data'].append({'excel_row': index, **vals})
                    row.clear()
    for diag in result['diagnostics']:
        rows = diag['data']
        status = [c for c in diag['columns'] if c.lower() in ('status', 'result', 'outcome', 'passed', 'pass')]
        diag['status_counts'] = dict(collections.Counter(str(row.get(c)) for row in rows for c in status))
        diag['nonpass'] = [row for row in rows if any(str(row.get(c)).strip().lower()
                          not in ('pass', 'passed', 'ok', 'true', 'success') for c in status)]
    result['unchanged'] = before == digest(path)
    return result

def main():
    jobs = json.loads((OUT / 'selected-workbooks.json').read_text(encoding='utf-8-sig'))
    ledger = [p for run in ('20261029-074441-a2fe3870', '20261029-085859-fbb9b367')
              for p in (DATE / 'RunLogs/BatchRefresh' / run).rglob('*') if p.is_file()]
    stamps = {str(p): digest(p) for p in ledger}
    results = []
    for i, job in enumerate(jobs):
        try:
            results.append(inspect(job))
        except Exception as exc:
            results.append({'id': job['Id'], 'error': repr(exc), 'unchanged': False})
        if (i+1) % 10 == 0:
            print(f'Inspected {i+1}/{len(jobs)}', flush=True)
    summary = {'workbooks': len(results), 'inspection_errors': [r for r in results if 'error' in r],
               'all_workbooks_unchanged': all(r['unchanged'] for r in results),
               'ledger_unchanged': all(digest(Path(p)) == h for p, h in stamps.items()),
               'tables': sum(len(r.get('tables', [])) for r in results),
               'diagnostic_tables': sum(len(r.get('diagnostics', [])) for r in results),
               'cell_error_count': sum(len(r.get('cell_errors', [])) for r in results),
               'nonpass_tables': [{'id': r['id'], 'table': d['name'], 'sheet': d['sheet'], 'ref': d['ref'],
                                   'counts': d['status_counts'], 'nonpass': d['nonpass']}
                                 for r in results for d in r.get('diagnostics', []) if d['nonpass']],
               'unloaded_check_connections': [{'id': r['id'], 'connections': r['unloaded_check_connections']}
                                             for r in results if r.get('unloaded_check_connections')],
               'broken_names': [{'id': r['id'], 'names': r['broken_defined_names']}
                                for r in results if r.get('broken_defined_names')]}
    (OUT / 'saved-checks.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
    (OUT / 'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k:v for k,v in summary.items() if k not in ('nonpass_tables', 'unloaded_check_connections', 'broken_names')}, indent=2))
    print('Nonpass tables:', len(summary['nonpass_tables']))
    print('Unloaded check connections:', len(summary['unloaded_check_connections']))
    print('Workbooks with broken names:', len(summary['broken_names']))

if __name__ == '__main__':
    main()
