"""Read only the three authorised BD upstream workbooks, without Excel."""

import importlib.util
import json
import sys
import collections
import datetime
import hashlib
import math
from pathlib import Path

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[2]
HELPER = ROOT / "outputs/jh-ry-ain-availability-20261021/inspect_saved_tables.py"
spec = importlib.util.spec_from_file_location("saved_tables_reader", HELPER)
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)
PREFIX = "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/"
FILES = ["StaffListMaster.xlsx", "Settings Data.xlsx", "AllocationByShiftAverage.xlsx"]


def main():
    selected = {
        "StaffListMaster.xlsx": {"Table_Masterlist", "ResourceContract_CHECK"},
        "Settings Data.xlsx": {"MaxAvailability", "ShiftDuration", "PermutationDimensions", "Table_DateList", "Table_Min_Date", "DateFrom"},
        "AllocationByShiftAverage.xlsx": {"ResourceShiftAllocation"},
    }
    books = {name: reader.read_book(PREFIX + name, selected[name]) for name in FILES}
    tables = {name: {t["name"]: t for t in book["tables"] if "data" in t} for name, book in books.items()}

    def serial_date(value):
        if not isinstance(value, (int, float)):
            return value
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)).isoformat()

    def located_rows(table):
        top = reader.bounds(table["ref"])[1]
        return [{"excel_row": top + offset + 1, **row} for offset, row in enumerate(table["data"])]

    master = tables["StaffListMaster.xlsx"]["Table_Masterlist"]
    workers = [r for r in located_rows(master) if r.get("Resource") == 75]
    names = {r["Name"] for r in workers if r.get("Role") == "RN"}
    allocation = tables["AllocationByShiftAverage.xlsx"]["ResourceShiftAllocation"]
    matches = [r for r in located_rows(allocation) if r.get("Name") in names and r.get("Role") == "RN"]
    positives = [r for r in matches if isinstance(r.get("ResShiftFTE"), (int, float)) and r["ResShiftFTE"] > 0]
    below_threshold = [r for r in positives if r["ResShiftFTE"] < 0.5]
    group = collections.defaultdict(list)
    for row in positives:
        group[(row.get("ShiftDate"), row.get("ShiftPeriod"))].append(row)
    grouped = [{"ShiftDate": date, "date_iso": serial_date(date), "ShiftPeriod": shift,
                "positive_rows": len(rows), "sum_ResShiftFTE": sum(r["ResShiftFTE"] for r in rows),
                "excel_rows": [r["excel_row"] for r in rows]} for (date, shift), rows in sorted(group.items())]
    setting_dates = [r["Date"] for r in tables["Settings Data.xlsx"]["Table_DateList"]["data"]]
    start, end = min(setting_dates), max(setting_dates)
    in_window = [r for r in positives if start <= r["ShiftDate"] <= end]
    grouped_in_window = [r for r in grouped if start <= r["ShiftDate"] <= end]
    period_lookup = {(r["Date"], r["Shifts"]): r["Period"]
                     for r in tables["Settings Data.xlsx"]["PermutationDimensions"]["data"] if r["RolesList"] == "RN"}
    for row in grouped:
        row["Period"] = period_lookup.get((row["ShiftDate"], row["ShiftPeriod"]))
    positive_periods = sorted(r["Period"] for r in grouped_in_window)
    period_runs = []
    for period in positive_periods:
        if period_runs and period == period_runs[-1][1] + 1:
            period_runs[-1][1] = period
        else:
            period_runs.append([period, period])
    shift_groups = collections.defaultdict(list)
    for row in grouped_in_window:
        shift_groups[row["ShiftPeriod"]].append(row)
    window_summary = {"start_date": serial_date(start), "end_date": serial_date(end),
                      "raw_positive_rows": len(in_window), "positive_shift_date_keys": len(grouped_in_window),
                      "sum_positive_ResShiftFTE": sum(r["ResShiftFTE"] for r in in_window),
                      "raw_positive_rows_below_0_5": sum(r["ResShiftFTE"] < 0.5 for r in in_window),
                      "grouped_positive_keys_below_0_5": sum(r["sum_ResShiftFTE"] < 0.5 for r in grouped_in_window),
                      "positive_periods": positive_periods, "positive_period_runs_inclusive": period_runs,
                      "ShiftPeriod_vs_IntervalAssociatedShift_mismatches": sum(r["ShiftPeriod"] != r["IntervalAssociatedShift"] for r in matches),
                      "by_shift": [{"ShiftPeriod": shift, "positive_shift_date_keys": len(rows),
                                    "sum_positive_ResShiftFTE": sum(r["sum_ResShiftFTE"] for r in rows)}
                                   for shift, rows in sorted(shift_groups.items())],
                      "out_of_window_keys": [r for r in grouped if not start <= r["ShiftDate"] <= end]}
    settings = []
    for table in tables["Settings Data.xlsx"].values():
        item = {k: v for k, v in table.items() if k != "data"}
        rows = table["data"]
        if table["name"] == "PermutationDimensions":
            item["roles"] = sorted({r.get("RolesList") for r in rows})
            item["shifts"] = sorted({r.get("Shifts") for r in rows})
            item["min_date"] = serial_date(min(r["Date"] for r in rows))
            item["max_date"] = serial_date(max(r["Date"] for r in rows))
            item["days"] = len({r["Date"] for r in rows})
            item["rn_period_rows"] = sum(r.get("RolesList") == "RN" for r in rows)
        else:
            item["data"] = located_rows(table)
        settings.append(item)
    evidence = {
        "scope": "Read-only saved values from three exact BD upstream workbooks; no Excel, refresh, recalculation, save or workbook write.",
        "files": [{"path": b["path"], "sha256_before": b["sha256"],
                   "sha256_after": hashlib.sha256((ROOT / b["path"]).read_bytes()).hexdigest(),
                   "unchanged_after_read": b["unchanged"]} for b in books.values()],
        "master": {"table": master["name"], "sheet": master["sheet"], "range": master["ref"], "resource75_rows": workers,
                   "derived_cap_for_current_settings_window": [
                       {"Resource": r["Resource"], "current_window_days_inclusive": end-start+1,
                        "current_window_fortnights": (end-start+1)/14,
                        "contracted_roster_hours": r["Contracted FN Hours"] * ((end-start+1)/14),
                        "shift_hours": tables["Settings Data.xlsx"]["ShiftDuration"]["data"][0]["ShiftDuration"],
                        "contracted_shift_equivalent": r["Contracted FN Hours"]*((end-start+1)/14)/tables["Settings Data.xlsx"]["ShiftDuration"]["data"][0]["ShiftDuration"],
                        "contracted_whole_shifts": math.floor(r["Contracted FN Hours"]*((end-start+1)/14)/tables["Settings Data.xlsx"]["ShiftDuration"]["data"][0]["ShiftDuration"]),
                        "effective_cap": min(tables["Settings Data.xlsx"]["MaxAvailability"]["data"][0]["MaxAvailability"], math.floor(r["Contracted FN Hours"]*((end-start+1)/14)/tables["Settings Data.xlsx"]["ShiftDuration"]["data"][0]["ShiftDuration"]))}
                       for r in workers],
                   "roster_dates": [{"Resource": r["Resource"], "start": serial_date(r.get("Roster Start")), "end": serial_date(r.get("Roster End"))} for r in workers],
                   "contract_checks": tables["StaffListMaster.xlsx"]["ResourceContract_CHECK"]},
        "settings": settings,
        "allocation": {"table": allocation["name"], "sheet": allocation["sheet"], "range": allocation["ref"],
                       "worker_match": "Exact Name and Role=RN from Resource75's saved master row",
                       "matching_rows": len(matches), "positive_rows": len(positives), "positive_shift_date_keys": len(grouped),
                       "sum_positive_ResShiftFTE": sum(r["ResShiftFTE"] for r in positives),
                       "positive_below_0_5_rows": len(below_threshold),
                       "sum_positive_below_0_5": sum(r["ResShiftFTE"] for r in below_threshold),
                       "current_settings_window": window_summary,
                       "positive_grouped_by_shift_date": grouped, "matching_rows_data": matches},
    }
    target = Path(__file__).with_name("upstream75-evidence.json")
    target.write_text(json.dumps(evidence, indent=2), encoding="utf-8")
    compact = {**evidence, "allocation": {k: v for k, v in evidence["allocation"].items() if k not in {"matching_rows_data"}}}
    print(json.dumps(compact, ensure_ascii=False))


if __name__ == "__main__":
    main()
