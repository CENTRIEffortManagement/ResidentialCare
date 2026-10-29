"""Recheck audit hashes and prepare a concise review from the inspected evidence."""
import collections
import json
import runpy
from pathlib import Path

scope = runpy.run_path(str(Path(__file__).with_name('inspect_saved_checks.py')))
OUT, DATE = scope['OUT'], scope['DATE']
jobs = json.loads((OUT/'selected-workbooks.json').read_text(encoding='utf-8-sig'))
books = json.loads((OUT/'saved-checks.json').read_text(encoding='utf-8'))
by_id = {j['Id']: j for j in jobs}
reinspected = []
for index, book in enumerate(books):
    if scope['digest']((DATE/book['path']).resolve()) != book['sha256']:
        books[index] = scope['inspect'](by_id[book['id']])
        reinspected.append(book['id'])
    books[index]['current_metadata'] = by_id[book['id']]['Metadata']['OutputStatus']
(OUT/'saved-checks.json').write_text(json.dumps(books, ensure_ascii=False, indent=2), encoding='utf-8')
severity = []
for book in books:
    for table in book['diagnostics']:
        for column in table['columns']:
            if column == 'Severity' or column.endswith('Status'):
                counts = dict(collections.Counter(str(r.get(column)) for r in table['data']))
                severity.append({'id':book['id'], 'table':table['name'], 'column':column, 'counts':counts})
(OUT/'all-check-statuses.json').write_text(json.dumps(severity, indent=2), encoding='utf-8')
print(json.dumps({'reinspected_after_external_changes':reinspected,
                  'all_inspected_files_stable_during_reads':all(b['unchanged'] for b in books),
                  'current_metadata':dict(collections.Counter(j['Metadata']['OutputStatus'] for j in jobs))}, indent=2))
