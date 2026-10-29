"""Summarize the saved audit and locate cell errors against loaded tables."""
import collections
import json
import runpy
from pathlib import Path
from zipfile import ZipFile

scope = runpy.run_path(str(Path(__file__).with_name('inspect_saved_checks.py')))
OUT, DATE, ET, S = (scope[k] for k in ('OUT', 'DATE', 'ET', 'S'))
books = json.loads((OUT / 'saved-checks.json').read_text(encoding='utf-8'))
totals = collections.Counter()
error_tables, samples, no_status, role_summary, exception_rows = [], [], [], [], []
for book in books:
    for d in book['diagnostics']:
        totals.update(d['status_counts'])
        if not d['status_counts'] and d['observed_rows']:
            no_status.append({'id': book['id'], 'name': d['name'], 'rows': d['observed_rows'], 'columns': d['columns']})
        if d['name'] == 'A1AllocationIdentity_EXCEPTIONS':
            exception_rows.append({'id': book['id'], 'rows': d['observed_rows'], 'columns': d['columns'],
                                   'data': [{k:v for k,v in r.items() if 'name' not in k.lower()} for r in d['data']]})
        if d['name'] in ('CapacityDiagnostics_SUMMARY', 'BFinalCap_CHECK') and d['nonpass']:
            role_summary.append({'id': book['id'], 'table': d['name'], 'failures': d['nonpass']})
    errmap = collections.Counter()
    for error in book['cell_errors']:
        x = scope['re'].fullmatch(r'([A-Z]+)(\d+)', error['cell'])
        col, row = scope['colnum'](x[1]), int(x[2])
        hits = []
        for table in book['tables']:
            left, top, right, bottom = scope['bounds'](table['ref'])
            if table['sheet'] == error['sheet'] and left <= col <= right and top <= row <= bottom:
                hits.append(table['name'])
        errmap[(error['sheet'], error['error'], ','.join(hits) or '(outside Excel tables)')] += 1
    if not errmap:
        continue
    error_tables.append({'id': book['id'], 'counts': [{'sheet': s, 'error': e, 'table': t, 'count': n}
                                                    for (s,e,t),n in errmap.items()]})
    with ZipFile((DATE / book['path']).resolve(), 'r') as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        wr = scope['rels'](z, 'xl/workbook.xml')
        for sheet in wb.find(S+'sheets'):
            name = sheet.attrib['name']
            targets = [e for e in book['cell_errors'] if e['sheet'] == name][:3]
            if not targets:
                continue
            coords = {e['cell'] for e in targets}
            with z.open(wr[sheet.attrib[scope['REL']+'id']]) as stream:
                for _, c in ET.iterparse(stream, events=('end',)):
                    if c.tag == S+'c' and c.attrib.get('r') in coords:
                        f, v = c.find(S+'f'), c.find(S+'v')
                        samples.append({'id': book['id'], 'sheet': name, 'cell': c.attrib['r'],
                                        'formula': f.text if f is not None else None,
                                        'error': v.text if v is not None else None})
                        coords.remove(c.attrib['r'])
                    if c.tag == S+'row':
                        c.clear()
                    if not coords:
                        break
result = {'status_counts': dict(totals), 'error_locations': error_tables, 'error_samples': samples,
          'nonempty_diagnostics_without_status': no_status, 'allocation_exceptions': exception_rows,
          'capacity_failures': role_summary}
(OUT / 'triage.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({k:result[k] for k in ('status_counts', 'error_locations', 'error_samples', 'nonempty_diagnostics_without_status')}, indent=2))
