"""Read saved values for the two specifically authorised BD/RN files."""
import collections
import importlib.util
import json
from pathlib import Path

spec = importlib.util.spec_from_file_location("saved_reader", "outputs/jh-ry-ain-availability-20261021/inspect_saved_tables.py")
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)

names = {"ResPeriodAvailabilityTABLE", "ResPeriodAllocationTABLE", "ResPeriodCapPrioritised", "ResPeriodWDTABLE",
         "ResourcePeriodTABLE_empty", "C__ResSingleShiftDayTABLE", "ResPeriodCappedAvailability_C__TABLE",
         "ResPeriodAvailabilityCapped_C__TABLE", "CapacityDiagnostics_SUMMARY", "CapacityDiagnostics_DETAILS", "Demand_Prepare"}
files = ["CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/RN/CapacityDistrib(A.1)-shifts.xlsx",
         "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/RN/CapacityDistrib(A.2)-shifts.xlsx"]
evidence, summary = [], []
for filename in files:
    book = reader.read_book(filename, names)
    result = {k: book[k] for k in ("path", "sha256", "unchanged")}
    overview = result.copy()
    result["tables"], overview["tables"] = [], []
    for table in book["tables"]:
        if "data" not in table:
            continue
        top = reader.bounds(table["ref"])[1]
        rows = [dict(excel_row=top + i + 1, **row) for i, row in enumerate(table["data"])
                if any(value is not None for value in row.values())]
        selected = rows if table["name"] == "CapacityDiagnostics_SUMMARY" else [r for r in rows if r.get("Resource") == 75]
        info = {k: v for k, v in table.items() if k != "data"}
        info["data"] = selected
        result["tables"].append(info)
        item = {k: info[k] for k in ("name", "sheet", "ref")}
        item["selected_rows"] = len(selected)
        if table["name"] == "CapacityDiagnostics_SUMMARY":
            item["nonpass"] = [r for r in rows if r.get("Status") != "Pass"]
        elif table["name"] == "Demand_Prepare":
            item["date_serials"] = sorted({r.get("Date") for r in rows})
            item["periods"] = sorted({r.get("Period") for r in rows})
            item["roles"] = sorted({r.get("Role") for r in rows})
        elif selected:
            item["numeric_sums"] = {key: sum(r[key] for r in selected if isinstance(r.get(key), (int, float)) and not isinstance(r[key], bool))
                                    for key in ("OriginalAvailability", "Availability", "AvailabilityCapped", "Allocation", "UnassignedAvail")
                                    if key in table["columns"]}
            item["flags"] = {key: [{"value": v, "count": n} for v, n in collections.Counter(r.get(key) for r in selected).items()]
                             for key in ("HasAllocationEvidence", "SpareCalculationAllowed", "SpareDayEligible", "SparePeriodSelected",
                                         "AllocationWorkday", "RosteredPeriodStatus", "SpareIssueReason", "Reason", "ResAv-C'")
                             if key in table["columns"]}
            counts = collections.Counter((r.get("Resource"), r.get("Period")) for r in selected)
            item["duplicate_resource_period_keys"] = [{"key": k, "count": n} for k, n in counts.items() if n > 1]
            item["periods"] = sorted({r["Period"] for r in selected if r.get("Period") is not None})
        overview["tables"].append(item)
    evidence.append(result)
    summary.append(overview)

folder = Path("outputs/bd-rn-resource75-20261021")
(folder / "a1_a2_saved_evidence.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding="utf-8")
(folder / "a1_a2_saved_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(summary, ensure_ascii=False))
