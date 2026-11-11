"""Read only the declared dimensions of the three configured roster sheets."""
from pathlib import Path
from zipfile import ZipFile
import json
import posixpath
import xml.etree.ElementTree as ET

repo = Path(__file__).resolve().parents[3]
run = repo / (repo / 'outputs/overtime-roster-checks/deployment-current.txt').read_text().strip()
ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
rid = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
rows = []
for unit in ('BD', 'TE', 'JH-RY'):
    path = repo / f'CLIENT/DATExx-Whiddon/UNITS/{unit}/1. Input/Allocation/Published Roster.xlsx'
    with ZipFile(path) as archive:
        book = ET.fromstring(archive.read('xl/workbook.xml'))
        rels = {x.attrib['Id']: x.attrib['Target'] for x in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        sheet = next(s for s in book.find(ns+'sheets') if s.attrib['name'] == 'Roster Data')
        target = rels[sheet.attrib[rid]]
        part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
        dimension = None
        with archive.open(part) as stream:
            for event, element in ET.iterparse(stream, events=('end',)):
                if element.tag == ns+'dimension':
                    dimension = element.attrib['ref']
                    break
        first_cells = []
        with archive.open(part) as stream:
            for event, element in ET.iterparse(stream, events=('end',)):
                if element.tag == ns+'row':
                    first_cells = [(c.attrib.get('t'), c.findtext(ns+'v'), ''.join(t.text or '' for t in c.iter(ns+'t'))) for c in element]
                    break
        shared_indices = [int(value) for kind, value, text in first_cells if kind == 's' and value is not None]
        strings = []
        if shared_indices:
            with archive.open('xl/sharedStrings.xml') as stream:
                for event, element in ET.iterparse(stream, events=('end',)):
                    if element.tag == ns+'si':
                        strings.append(''.join(t.text or '' for t in element.iter(ns+'t')))
                        element.clear()
                        if len(strings) > max(shared_indices):
                            break
        headers = [strings[int(value)] if kind == 's' else text if kind == 'inlineStr' else value for kind, value, text in first_cells]
        rows.append({'Unit':unit, 'Workbook':path.relative_to(repo).as_posix(), 'RosterSheetDimension':dimension,
                     'RosterSheetXMLBytes':archive.getinfo(part).file_size, 'WorkbookBytes':path.stat().st_size,
                     'RosterHeaders':headers})
(run / 'roster-sheet-dimensions.json').write_text(json.dumps(rows, indent=2)+'\n')
print(json.dumps(rows, indent=2))
