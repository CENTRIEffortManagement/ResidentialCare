"""Compare authorized saved-workbook extracts with their adjacent M sources."""

import difflib
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
BASE = ROOT / "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations"
PAIRS = {
    "RN-A1": "RN/CapacityDistrib(A.1)-shifts.xlsx",
    "RN-A2": "RN/CapacityDistrib(A.2)-shifts.xlsx",
    "RN-B": "RN/CapacityDistrib(B)-shifts.xlsx",
    "StaffListMaster": "StaffListMaster.xlsx",
    "Settings": "Settings Data.xlsx",
    "Allocation": "AllocationByShiftAverage.xlsx",
}


def normalize(value):
    value = value.replace("\r\n", "\n").replace("\r", "\n").lstrip("\ufeff")
    section = re.search(r"(?m)^section\s+", value)
    return value[section.start():].strip() if section else value.strip()


def definitions(value):
    matches = list(re.finditer(r'(?m)^shared\s+(#"(?:[^"]|"")*"|[\w]+)\s*=', value))
    result = {}
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(value)
        body = value[match.start():end]
        body = "\n".join(line.strip() for line in body.splitlines()
                         if line.strip() and not line.lstrip().startswith("//"))
        result[match.group(1)] = body
    return result


report = []
for label, relative in PAIRS.items():
    source = BASE / (relative + "_PowerQuery.m")
    embedded = OUT / (label + ".embedded.m")
    actual = normalize(embedded.read_text(encoding="utf-8-sig"))
    if not source.exists():
        report.append({"label": label, "source_missing": True})
        continue
    expected = normalize(source.read_text(encoding="utf-8-sig"))
    actual_queries, expected_queries = definitions(actual), definitions(expected)
    different = [name for name in expected_queries.keys() & actual_queries.keys()
                 if expected_queries[name] != actual_queries[name]]
    delta = list(difflib.unified_diff(expected.splitlines(), actual.splitlines(),
                                    fromfile=str(source.relative_to(ROOT)),
                                    tofile=str(embedded.relative_to(ROOT)), n=2))
    (OUT / (label + ".source-vs-workbook.diff.txt")).write_text("\n".join(delta), encoding="utf-8")
    report.append({
        "label": label,
        "source": str(source.relative_to(ROOT)),
        "embedded": str(embedded.relative_to(ROOT)),
        "exact_section_match": expected == actual,
        "source_queries": len(expected_queries),
        "embedded_queries": len(actual_queries),
        "only_in_source": sorted(expected_queries.keys() - actual_queries.keys()),
        "only_in_workbook": sorted(actual_queries.keys() - expected_queries.keys()),
        "different_queries": sorted(different),
        "diff_lines": len(delta),
    })
(OUT / "source-parity.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
print(json.dumps(report, indent=2))
