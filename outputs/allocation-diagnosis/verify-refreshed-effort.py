"""Check the saved downstream output grain after the authorized refresh."""
from pathlib import Path
from collections import Counter
import json, math
import openpyxl

ROOT=Path(__file__).resolve().parents[2]
BASE=ROOT/'CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations'
OUT=Path(__file__).resolve().parent

def output_rows(file, columns):
    wb=openpyxl.load_workbook(BASE/file,read_only=True,data_only=True,keep_links=False)
    try:
        for ws in wb:
            values=ws.iter_rows(values_only=True)
            headers=next(values,())
            if set(columns).issubset(headers):
                indices=[headers.index(c) for c in columns]
                return [dict(zip(columns,[row[i] for i in indices])) for row in values if any(row[i] is not None for i in indices)]
    finally:
        wb.close()
    raise AssertionError('Output columns not found: '+file)

allocation=output_rows('Allocation.xlsx',['ShiftDate','ShiftPeriod','Role','Attribute','RoleShiftEffort','DurationOfShifts','RoleShiftFTE'])
effort=output_rows('Effort.xlsx',['Facility','Role','Date','Shift','Demand','Capacity','CapacityX','Allocation','DATESHIFT'])
key_a=lambda r:(str(r['ShiftDate']),r['Role'],r['ShiftPeriod'])
key_e=lambda r:(str(r['Date']),r['Role'],r['Shift'])
ac=Counter(map(key_a,allocation))
ec=Counter((r['Facility'],*key_e(r)) for r in effort)
assert len(allocation)==len(ac)==252, 'Allocation grain or calendar count incorrect'
assert len(effort)==len(ec)==252, 'Effort grain or calendar count incorrect'
assert all(r['Attribute'] in ('IntervalStart',None) for r in allocation)
lookup={key_a(r):r for r in allocation}
assert set(lookup)==set(map(key_e,effort)), 'Date/role/shift coverage differs'
assert all(math.isfinite(r['Allocation']) and math.isclose(r['Allocation'],lookup[key_e(r)]['RoleShiftFTE'],rel_tol=1e-10,abs_tol=1e-10) for r in effort), 'Allocation values do not reconcile'
example=[r for r in effort if r['Role']=='AIN' and r['Shift']=='AM' and str(r['Date']).startswith('2026-07-24')]
assert len(example)==1
assert math.isclose(example[0]['Allocation'],18.65328177,abs_tol=1e-7), 'User example differs'
result={'status':'passed','allocation_rows':len(allocation),'effort_rows':len(effort),'duplicate_allocation_keys':sum(v>1 for v in ac.values()),'duplicate_effort_keys':sum(v>1 for v in ec.values()),'allocation_values_reconciled':True,'effort_example':example[0],'files':[{'path':str((BASE/n).relative_to(ROOT)).replace('\\','/'),'modified':(BASE/n).stat().st_mtime} for n in ['Allocation.xlsx','Effort.xlsx']]}
(OUT/'refreshed-effort-verification.json').write_text(json.dumps(result,indent=2,default=str))
print(json.dumps(result,indent=2,default=str))
