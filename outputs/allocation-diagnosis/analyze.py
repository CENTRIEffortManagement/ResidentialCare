from pathlib import Path
from collections import Counter, defaultdict
import json, math

P = Path(__file__).resolve().parent
load = lambda name: json.loads((P / (name + '.xlsx.tables.json')).read_text())
s,a,b,e = [load(n) for n in ['Settings Data','AllocationByShiftAverage','Allocation','Effort']]
def profile(rows, keys):
    counts=Counter(tuple(r[k] for k in keys) for r in rows)
    return {'rows':len(rows),'keys':len(counts),'multiplicities':dict(Counter(counts.values())), 'min_date':min(str(r.get('Date',r.get('ShiftDate'))) for r in rows),'max_date':max(str(r.get('Date',r.get('ShiftDate'))) for r in rows)}
result={}
for label,rows,keys in [('upstream',a['Table_RoleShiftIntervalAllocation'],['ShiftDate','Role','ShiftPeriod']),('allocation',b['Allocation'],['ShiftDate','Role','ShiftPeriod']),('effort',e['EffortAllMatrixAG1_1D'],['Facility','Role','Date','Shift']),('demand',e['IMPORT_DemandCorrected'],['Date','Role','ShiftPeriod'])]:
    result[label]=profile(rows,keys)
result['shift_period']=s['ShiftPeriod']
key=('2026-07-20 00:00:00','AIN','AM')
result['allocation_example']=[r for r in b['Allocation'] if tuple(r[k] for k in ['ShiftDate','Role','ShiftPeriod'])==key]
result['effort_example']=[r for r in e['EffortAllMatrixAG1_1D'] if tuple(r[k] for k in ['Date','Role','Shift'])==key]
durations={(r['Role'],r['ShiftPeriod']):r['DurationOfShifts'] for r in s['ShiftPeriod']}
raw=a['ResourceShiftAllocationY']
totals=defaultdict(float)
for r in raw:
    totals[tuple(r[k] for k in ['ShiftDate','Role','ShiftPeriod','Attribute'])]+=r['ResEffectiveIntervalEffort']
result['aggregation_max_error']=max(abs(r['RoleShiftEffort']-totals[tuple(r[k] for k in ['ShiftDate','Role','ShiftPeriod','Attribute'])]) for r in a['Table_RoleShiftIntervalAllocation'])
result['fte_formula_max_error']=max(abs(r['RoleShiftFTE']-r['RoleShiftEffort']*24/r['DurationOfShifts']) for r in a['Table_RoleShiftIntervalAllocation'])
result['zero_staff_nonzero_effort']=sum(r['StaffCount']==0 and r['ResEffectiveIntervalEffort']!=0 for r in raw)
result['attributes']=dict(Counter(r['Attribute'] for r in raw))
result['wrong_denominator_rows']=sum(r['DurationOfShifts']!=durations[(r['Role'],r['ShiftPeriod'])] for r in b['Allocation'] if r['DurationOfShifts'] is not None)
result['null_padding_rows']=sum(r['Attribute'] is None for r in b['Allocation'])
groups=defaultdict(list)
for r in e['EffortAllMatrixAG1_1D']: groups[(r['Date'],r['Role'],r['Shift'])].append(r)
result['demand_capacity_invariant_groups']=sum(all(len({r[col] for r in rows})==1 for col in ['Demand','Capacity','CapacityMaxHC','CapacityX']) for rows in groups.values())
result['sum_inflation']={col:{'saved_sum':sum(r[col] for r in e['EffortAllMatrixAG1_1D']),'one_per_key_sum':sum(rows[0][col] for rows in groups.values())} for col in ['Demand','Capacity']}
result['example_start_fte']=totals[(*key,'IntervalStart')]*24/durations[(key[1],key[2])]
result['example_end_fte']=totals[(*key,'IntervalEnd')]*24/durations[(key[1],key[2])]
result['example_fte_sum']=sum(r['Allocation'] for r in result['effort_example'])
result['non_six_keys']=[{'date':k[0],'role':k[1],'shift':k[2],'count':len(v)} for k,v in groups.items() if len(v)!=6]
result['settings_keys_unique']=len(durations)==len(s['ShiftPeriod'])
# Equality of the saved upstream rows and Allocation, after the calendar restriction.
calendar={tuple(r[k] for k in ['Date','RolesList','Shifts']) for r in s['PermutationDimensions']}
up=[r for r in a['Table_RoleShiftIntervalAllocation'] if tuple(r[k] for k in ['ShiftDate','Role','ShiftPeriod']) in calendar]
sig=lambda r: tuple(r.get(k) for k in ['ShiftDate','Role','ShiftPeriod','Attribute','RoleShiftEffort','DurationOfShifts','RoleShiftFTE'])
result['allocation_equals_upstream_nonpadding']=Counter(map(sig,up))==Counter(sig(r) for r in b['Allocation'] if r['Attribute'] is not None)
result['effort_allocation_values_match']=all(all(math.isclose(x,y,rel_tol=1e-12,abs_tol=1e-12) for x,y in zip(sorted(r['Allocation'] for r in rows),sorted(r['RoleShiftFTE'] for r in b['Allocation'] if (r['ShiftDate'],r['Role'],r['ShiftPeriod'])==k))) for k,rows in groups.items())
result['start_end_effort_different_keys']=sum(not math.isclose(v,totals.get((d,role,shift,'IntervalEnd'),0),abs_tol=1e-10) for (d,role,shift,attr),v in totals.items() if attr=='IntervalStart' and (d,role,shift) in calendar)
(P/'evidence.json').write_text(json.dumps(result,indent=2))
print(json.dumps(result,indent=2))
