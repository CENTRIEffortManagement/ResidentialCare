"""Prepare a review patch only; the approved source is not changed by this script."""
from pathlib import Path
import difflib
ROOT=Path(__file__).resolve().parents[2]
OUT=Path(__file__).resolve().parent
TARGET=Path('CLIENT/DATExx-Whiddon/UNITS/Unit1/2. Calculations/AllocationByShiftAverage.xlsx_PowerQuery.m')
original=(ROOT/TARGET).read_text(encoding='utf-8-sig')
replacement=(OUT/'allocation-query-replacement.txt').read_text()
start=original.index('shared RoleShiftAllocation =')
end=original.index('shared RoleShiftEffortFTE =',start)
candidate=original[:start]+replacement+'\n'+original[end:]
(OUT/'AllocationByShiftAverage.proposed.txt').write_text(candidate,encoding='utf-8')
(OUT/'allocation-fix.patch').write_text(''.join(difflib.unified_diff(original.splitlines(True),candidate.splitlines(True),fromfile=str(TARGET).replace('\\','/'),tofile=str(TARGET).replace('\\','/'))),encoding='utf-8')
print('Prepared proposed source and review patch; production M unchanged.')
