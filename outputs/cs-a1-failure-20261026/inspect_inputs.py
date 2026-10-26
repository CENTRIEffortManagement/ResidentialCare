"""Read the exact CS A.1 saved inputs; never open Excel or write workbooks."""
import collections
import datetime
import hashlib
import importlib.util
import json
import math
import re
import sys
from pathlib import Path

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location(
    "saved_reader", ROOT / "outputs/jh-ry-ain-availability-20261021/inspect_saved_tables.py"
)
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
BASE = "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/"
TARGETS = {
    "CS/CapacityDistrib(A.1)-shifts.xlsx": {
        "RolePathTABLE", "FilePathUrl", "CapacityDiagnostics_SUMMARY", "CapacityDiagnostics_DETAILS"
    },
    "StaffListMaster.xlsx": {"Table_Masterlist", "ResourceContract_CHECK"},
    "Capacity-ShiftAvailability.xlsx": {"ResDayShift", "ResDayShift_CHECK", "Availability_StaffList"},
    "AllocationByShiftAverage.xlsx": {"ResourceShiftAllocation"},
    "Demand.xlsx": {"ShiftDemandHCAverageANACC"},
    "Settings Data.xlsx": {"PermutationDimensions", "DateFrom", "MaxAvailability", "ShiftDuration"},
}

def active(table):
    return [r for r in table.get("data", []) if any(v is not None for v in r.values())]

def isnum(x):
    return isinstance(x, (float, int)) and not isinstance(x, bool) and math.isfinite(x)

books = {}
summary = {}
for suffix, selected in TARGETS.items():
    book = reader.read_book(BASE + suffix, selected)
    books[suffix] = book
    info = {"path": book["path"], "sha256": book["sha256"], "unchanged": book["unchanged"],
            "tables": []}
    for table in book["tables"]:
        item = {k:v for k,v in table.items() if k != "data"}
        if "data" in table:
            rows = active(table)
            item["actual_rows"] = len(rows)
            rolecol = "RolesList" if table["name"] == "PermutationDimensions" else "Role"
            if rolecol in table["columns"]:
                item["by_role"] = dict(collections.Counter(r.get(rolecol) for r in rows))
            if table["name"].endswith("CHECK") or table["name"].endswith("SUMMARY"):
                item["failures"] = [r for r in rows if r.get("Status") not in ("Pass", "Warning")]
            if table["name"] in {"RolePathTABLE", "FilePathUrl", "DateFrom", "MaxAvailability", "ShiftDuration"}:
                item["values"] = rows
        info["tables"].append(item)
    summary[suffix] = info
    print(json.dumps({"inventory": suffix, "selected": [x for x in info["tables"] if "actual_rows" in x],
                      "unchanged": book["unchanged"]}, default=str), flush=True)

def get(suffix, table):
    return active(next(t for t in books[suffix]["tables"] if t["name"] == table))

masters = get("StaffListMaster.xlsx", "Table_Masterlist")
alloc = get("AllocationByShiftAverage.xlsx", "ResourceShiftAllocation")
avail = get("Capacity-ShiftAvailability.xlsx", "ResDayShift")
calendar = get("Settings Data.xlsx", "PermutationDimensions")
demand = get("Demand.xlsx", "ShiftDemandHCAverageANACC")
datefrom = get("Settings Data.xlsx", "DateFrom")[0]["DateFrom"]
repro = {}
for role in ("AIN", "CS"):
    workers = [r for r in masters if r.get("Role") == role]
    keys = collections.defaultdict(list)
    resource_counts = collections.Counter(r.get("Resource") for r in workers)
    for row in workers:
        keys[(row.get("Name"), row.get("Role"))].append(row)
    period_keys = collections.Counter((r.get("Date"), r.get("Shifts"), r.get("RolesList"))
                                      for r in calendar if r.get("RolesList") == role)
    ar = [r for r in alloc if r.get("Role") == role]
    scoped = [r for r in ar if isnum(r.get("ShiftDate")) and r["ShiftDate"] >= datefrom]
    usable = [r for r in scoped if len(keys[(r.get("Name"), role)]) == 1
              and resource_counts[keys[(r.get("Name"), role)][0].get("Resource")] == 1
              and isnum(r.get("ResShiftFTE"))]
    vr = [r for r in avail if r.get("Role") == role]
    matched_av = [r for r in vr if len(keys[(r.get("Name"), role)]) == 1]
    grouped = collections.defaultdict(float)
    for r in usable:
        resource = keys[(r.get("Name"), role)][0]["Resource"]
        grouped[(resource, r.get("ShiftDate"), r.get("IntervalAssociatedShift"))] += r["ResShiftFTE"]
    cells = {k:v for k,v in grouped.items() if v > 0.2}
    repro[role] = {
        "master_workers": len(workers), "allocation_rows": len(ar), "allocation_after_date": len(scoped),
        "allocation_mapped_usable": len(usable), "allocation_unmatched_names": len({
            r.get("Name") for r in scoped if len(keys[(r.get("Name"), role)]) != 1}),
        "final_allocation_cells_over_0_2": len(cells),
        "allocation_cells_without_calendar": sum(period_keys[(day, shift, role)] != 1
                                                  for _, day, shift in cells),
        "availability_rows": len(vr), "availability_identity_matches": len(matched_av),
        "availability_unmatched_names": len({r.get("Name") for r in vr
                                              if len(keys[(r.get("Name"), role)]) != 1}),
        "calendar_rows": sum(r.get("RolesList") == role for r in calendar),
        "demand_rows": sum(r.get("Role") == role for r in demand),
        "duplicate_master_name_role_keys": sum(len(v)>1 for v in keys.values()),
        "duplicate_master_resource_keys": sum(n>1 for n in resource_counts.values()),
        "unmatched_allocation_sample": sorted({r.get("Name") for r in scoped
                                               if len(keys[(r.get("Name"), role)]) != 1}, key=str)[:8],
    }

gate = ROOT / "outputs/cs-a1-failure-20261026/source-gate/workbook-extracted.m"
source = ROOT / "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/CS/CapacityDistrib(A.1)-shifts.xlsx_PowerQuery.m"
def normalize(p):
    text = p.read_text(encoding="utf-8-sig").replace("\r\n", "\n")
    return text[text.index("section "):].strip()
evidence = {"scope": "Saved values only; no Excel, recalculation, refresh, save, or workbook writes.",
            "source_gate_equal": normalize(gate) == normalize(source),
            "books": summary, "join_reproduction": repro}
path = Path(__file__).with_name("input-evidence.json")
path.write_text(json.dumps(evidence, indent=2, default=str), encoding="utf-8")
print(json.dumps({"source_gate_equal": evidence["source_gate_equal"], "join_reproduction": repro}, default=str))
