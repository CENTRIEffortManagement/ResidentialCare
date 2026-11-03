"""Read-only check of the 81 historical cap rows against current saved outputs."""
import collections
import json
import math
import runpy
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
read_book = runpy.run_path(str(ROOT/'outputs/jh-ry-ain-availability-20261021/inspect_saved_tables.py'))['read_book']
prior = json.loads((ROOT/'outputs/post-run-checks-20261029/saved-checks.json').read_text(encoding='utf-8'))
targets = {}
for b in prior:
    if b['id'] in ('BD/Role:AIN/3','BD/Role:CS/3','JH-RY/Role:CS/3','TE/Role:CS/3'):
        unit, role = b['id'].split('/')[0], b['id'].split(':')[1].split('/')[0]
        check = next(t for t in b['diagnostics'] if t['name'] == 'BFinalCap_CHECK')
        targets[(unit,role)] = [r['Resource'] for r in check['data'] if r.get('Status') == 'Fail']

evidence = []
def book(path, names):
    result = read_book('CLIENT/DATExx-Whiddon/'+path, set(names))
    if not result['unchanged']:
        raise RuntimeError('Workbook changed during inspection: '+path)
    evidence.append({'path':path,'sha256':result['sha256'],'unchanged':result['unchanged']})
    return {t['name']:[r for r in t['data'] if any(v is not None for v in r.values())]
            for t in result['tables'] if 'data' in t}

def numeric(v):
    return isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v)

def equal(a,b):
    return numeric(a) and numeric(b) and abs(a-b) < 1e-8

unit_books = {u:book(f'UNITS/{u}/2. Calculations/Effort.xlsx',
                    ['ResShiftAllocation','EffortAllMatrixAG1_1D']) for u in ('BD','TE','JH-RY')}
org = book('2. Calculations/E-O-I/Effort-All.xlsx',['ResShiftAllocation','EffortAllMatrixAG1_1D_2'])
outcomes = book('2. Calculations/E-O-I/EffortOutcomes.xlsx',['EffortOutcomesAG1_1DayShift'])
ineff = book('2. Calculations/E-O-I/Inefficiencies.xlsx',
             ['EffortOutcomesAG1_1DayShiftAB','InefficienciesAG1_1Level3210'])

def reskey(r,unit=False):
    return (r['Facility'], r['Role'], r['Resource'], r['Period'])
org_res = collections.defaultdict(float)
for r in org['ResShiftAllocation']:
    if numeric(r.get('Effort')):
        org_res[reskey(r)] += r['Effort']

def matrixkey(r, ineff=False):
    return (r.get('Facility'),r.get('L1.1') if ineff else r.get('Role'),r.get('Date'),
            r.get('L1.2') if ineff else r.get('Shift'))
def unique_index(rows, field, is_ineff=False):
    result = {}
    for row in rows:
        key = matrixkey(row,is_ineff)
        if key in result:
            raise RuntimeError('Duplicate reporting key: '+repr(key))
        result[key] = row.get(field)
    return result
indices = {
    'EffortAll':unique_index(org['EffortAllMatrixAG1_1D_2'],'Allocation'),
    'Outcomes':unique_index(outcomes['EffortOutcomesAG1_1DayShift'],'Allocation'),
    'InefficienciesInput':unique_index(ineff['EffortOutcomesAG1_1DayShiftAB'],'A',True),
    'InefficienciesFinal':unique_index([r for r in ineff['InefficienciesAG1_1Level3210']
                           if r.get('Level') == '3-Pool' and r.get('Version') == 'X'],'A')
}
groups = []
all_failures = []
for (unit,role), resources in targets.items():
    resset = set(resources)
    allocations = [r for r in unit_books[unit]['ResShiftAllocation'] if r['Role']==role and r['Resource'] in resset]
    present = {r['Resource'] for r in allocations if numeric(r.get('ResShiftFTE')) and r['ResShiftFTE']>0}
    by_res = collections.defaultdict(float)
    for r in allocations:
        if numeric(r.get('ResShiftFTE')):
            by_res[reskey(r)] += r['ResShiftFTE']
    missing_resource_period = [dict(zip(('Facility','Role','Resource','Period'),k)) for k,v in by_res.items()
                               if not equal(v,org_res.get(k))]
    affected = {(unit,role,r['TimeDate'],r['ShiftPeriod']) for r in allocations if numeric(r.get('ResShiftFTE')) and r['ResShiftFTE']>0}
    unit_matrix = unique_index(unit_books[unit]['EffortAllMatrixAG1_1D'],'Allocation')
    comparisons = {}
    for stage,index in indices.items():
        differences = [{'key':k,'UnitAllocation':unit_matrix.get(k),'DownstreamAllocation':index.get(k)}
                       for k in affected if not equal(unit_matrix.get(k),index.get(k))]
        comparisons[stage] = {'keys':len(affected),'differences':differences}
    current = book(f'UNITS/{unit}/2. Calculations/{role}/CapacityDistrib(B)-shifts.xlsx',
                   ['BFinalCap_CHECK','C___TABLE_B'])
    current_checks = {r['Resource']:r for r in current['BFinalCap_CHECK']}
    current_target = [current_checks.get(r) for r in resources]
    group = {'unit':unit,'role':role,'original_cap_rows':len(resources),
             'resources_in_saved_unit_allocation':len(present),'missing_resources':sorted(resset-present),
             'allocation_rows':len(allocations),'allocation_FTE_sum':sum(by_res.values()),
             'resource_periods':len(by_res),'organization_resource_period_differences':missing_resource_period,
             'downstream_allocation_comparisons':comparisons,
             'current_statuses':dict(collections.Counter(r.get('Status') if r else 'Missing' for r in current_target)),
             'current_origins':dict(collections.Counter(r.get('Origin') if r else 'Missing' for r in current_target))}
    groups.append(group)
    if group['missing_resources'] or missing_resource_period or any(x['differences'] for x in comparisons.values()):
        all_failures.append(unit+'/'+role)

report = {'scope':'81 cap-check rows recorded on 29 October; trace against current saved workbooks on 31 October',
          'original_cap_rows':sum(len(v) for v in targets.values()),'groups':groups,
          'groups_with_lineage_differences':all_failures,'evidence':evidence,
          'limitations':['Current files have changed since the original audit.',
                         'Cap excess is a Resource/roster measure; inefficiency classifications use aggregated demand/capacity/allocation.',
                         'No workbook was opened in Excel, refreshed, saved or changed.']}
(OUT/'trace.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
for g in groups:
    print(json.dumps({k:v for k,v in g.items() if k not in ('downstream_allocation_comparisons','organization_resource_period_differences')},ensure_ascii=False))
    print(json.dumps({'resource_period_differences':len(g['organization_resource_period_differences']),
                      'downstream_differences':{k:len(v['differences']) for k,v in g['downstream_allocation_comparisons'].items()}}))
print('Groups with differences:',all_failures)
