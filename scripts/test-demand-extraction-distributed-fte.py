"""Independent demand-adapter fixtures and optional saved-table reconciliation.
Run from the repo root: python scripts/test-demand-extraction-distributed-fte.py
Add --saved-workbooks to read the two exact input workbooks without changing them.
These checks do not execute Power Query or refresh Excel.
"""
import copy
import datetime as dt
import math
from pathlib import Path
import posixpath
import re
import sys
import unittest
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
UNIT = ROOT / "CLIENT/DATExx-Whiddon/UNITS/Unit1"
SOURCE = UNIT / "1. Input/2-DemandExtract.xlsx_PowerQuery.m"
ROLES = {"RN": "RN", "REGISTERED NURSE": "RN", "AIN": "AIN", "ASSISTANT IN NURSING": "AIN",
         "EN": "AINC4", "ENROLLED NURSE": "AINC4", "AINC4": "AINC4"}
RETAINED_ROLES = ("RN", "AIN", "AINC4")
# Explicit test configuration; saved-workbook mode reads the actual setting.
FIXTURE_HOURS = 7.6
SHIFTS = ("AM", "PM", "NIGHT")
COLUMNS = ["Date", "Day", "Shift", "Period", "Role", "StartTime", "EndTime",
           "Unit", "Facility", "DemandFTE", "DemandHRS", "DurationOfShifts"]


def require(condition, message):
    if not condition:
        raise ValueError(message)


def finite(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def standard_hours(rows):
    require(len(rows) == 1 and "ShiftDuration" in rows[0], "single ShiftDuration setting")
    hours = rows[0]["ShiftDuration"]
    require(finite(hours) and hours > 0, "positive finite ShiftDuration")
    return hours


def role(row):
    return ROLES.get(row["Role"].strip().upper())


def pattern_from_publication(allocation, profile, checks, retained_roles=RETAINED_ROLES,
                             shift_duration=FIXTURE_HOURS):
    require(finite(shift_duration) and shift_duration > 0, "standard hours")
    require(all(finite(r.get("ShiftDuration")) and r["ShiftDuration"] > 0
                and abs(r["ShiftDuration"] - shift_duration) <= 1e-7 / 60 for r in allocation + profile),
            "saved FTE basis")
    require(checks and all(r["Severity"] in ("Pass", "Warning") for r in checks), "upstream checks")
    def select(rows):
        return [r for r in rows if r["Facility"] == "BD" and role(r) in retained_roles]
    allocation, profile = select(allocation), select(profile)
    key = lambda r: (role(r), r["FortnightDayIndex"], "NIGHT" if r["Shift"] == "NS" else r["Shift"])
    expected = {(name, day, shift) for name in retained_roles for day in range(1, 15) for shift in SHIFTS}
    require(len(profile) == len(expected) and {key(r) for r in profile} == expected, "profile coverage")
    require(len(allocation) == len({key(r) for r in allocation}), "duplicate allocation")
    by_key = {key(r): r for r in allocation}
    require(set(by_key) <= {key(r) for r in profile}, "orphan allocation")
    result = {}
    for p in profile:
        k = key(p)
        require(k[1] in range(1, 15) and k[2] in SHIFTS, "pattern key")
        require(p["FortnightDayIndex"] == (p["FortnightWeek"] - 1) * 7 + p["DayOfWeek"], "day alignment")
        require(p["HistoricalCoverageStatus"] == "PASS", "history coverage")
        require(p["ProfileAlignmentStatus"] in ("PASS", "PASS ZERO"), "profile status")
        require(finite(p["RedistributedRosterFTE"]) and p["RedistributedRosterFTE"] >= 0, "profile FTE")
        a = by_key.get(k)
        require(p["RedistributionMatchCount"] == int(a is not None), "match count")
        if a is None:
            require(p["ProfileAlignmentStatus"] == "PASS ZERO" and p["HistoricalRosterFTE"] == 0
                    and p["RedistributedRosterFTE"] == 0, "unproven zero")
        else:
            require(a["Week No"] == p["Week No"], "week mismatch")
            require(finite(a["FTE"]) and a["FTE"] >= 0, "allocation FTE")
            require(abs(a["FTE"] - p["RedistributedRosterFTE"]) <= 1e-7, "paired FTE mismatch")
            require(a["Direct Care %"] == p["Direct Care %"], "direct-care mismatch")
        result[k] = p["RedistributedRosterFTE"]
    return result


def build_demand(pattern, dates, spans, retained_roles=RETAINED_ROLES, shift_duration=FIXTURE_HOURS):
    require(finite(shift_duration) and shift_duration > 0, "standard hours")
    require(len(dates) == 28 and dates[0].weekday() == 0 and
            dates == [dates[0] + dt.timedelta(days=i) for i in range(28)], "calendar")
    require(len(pattern) == len(retained_roles) * 14 * 3, "pattern coverage")
    output = []
    for offset, date in enumerate(dates):
        for shift_index, shift in enumerate(SHIFTS):
            for name in retained_roles:
                start, end, hours = spans[name, shift]
                start_time = dt.datetime.combine(date, start)
                end_time = dt.datetime.combine(date + dt.timedelta(days=shift == "NIGHT"), end)
                require(finite(hours) and hours > 0 and
                        abs((end_time - start_time).total_seconds() / 3600 - hours) <= 1e-7, "duration")
                fte = pattern[name, offset % 14 + 1, shift]
                require(finite(fte) and fte >= 0, "invalid FTE")
                output.append(dict(zip(COLUMNS, [date, offset + 1, shift, offset * 3 + shift_index + 1,
                    name, start_time, end_time, "BD", "BD", fte * shift_duration / hours, fte * shift_duration, hours])))
    return output


def fixture(third_role="EN", shift_duration=FIXTURE_HOURS):
    allocation, profile = [], []
    for name in RETAINED_ROLES:
        for day in range(1, 15):
            for shift in ("AM", "PM", "NS"):
                zero = day == 3 and shift == "PM"
                # Deliberately unequal weeks, fractional FTE and role-dependent amounts.
                value = 0 if zero else (1 if day <= 7 else 3) * (0.5 if name == "AIN" else 1.25)
                p = dict(Facility="BD", Role=third_role if name == "AINC4" else name, FortnightDayIndex=day,
                         FortnightWeek=(day - 1) // 7 + 1, DayOfWeek=(day - 1) % 7 + 1,
                         Shift=shift, **{"Week No": 30 + (day - 1) // 7, "Direct Care %": 0.75})
                p.update(ShiftDuration=shift_duration, HistoricalCoverageStatus="PASS", ProfileAlignmentStatus="PASS ZERO" if zero else "PASS",
                         HistoricalRosterFTE=value, RedistributedRosterFTE=value,
                         RedistributionMatchCount=0 if zero else 1)
                profile.append(p)
                if not zero:
                    allocation.append(dict(p, FTE=value))
    return allocation, profile, [dict(Severity="Pass")]


DATES = [dt.date(2026, 7, 20) + dt.timedelta(days=i) for i in range(28)]
SPANS = {(name, shift): span for name in RETAINED_ROLES for shift, span in {
    "AM": (dt.time(6), dt.time(14, 15), 8.25),
    "PM": (dt.time(14), dt.time(21, 15), 7.25),
    "NIGHT": (dt.time(21), dt.time(6), 9)
}.items()}


class DemandTests(unittest.TestCase):
    def test_unequal_weeks_repeat_without_rescaling(self):
        pattern = pattern_from_publication(*fixture())
        rows = build_demand(pattern, DATES, SPANS)
        self.assertEqual(len(rows), 252)
        self.assertEqual(len({r["Period"] for r in rows}), 84)
        by_key = {(r["Role"], r["Day"], r["Shift"]): r for r in rows}
        for r in rows:
            self.assertEqual(list(r), COLUMNS)
            self.assertAlmostEqual(r["DemandFTE"] * r["DurationOfShifts"], r["DemandHRS"])
            if r["Day"] <= 14:
                self.assertEqual(r["DemandHRS"], by_key[r["Role"], r["Day"] + 14, r["Shift"]]["DemandHRS"])
        self.assertEqual(by_key["RN", 8, "AM"]["DemandHRS"], 3 * by_key["RN", 1, "AM"]["DemandHRS"])
        self.assertEqual(by_key["RN", 3, "PM"]["DemandFTE"], 0)
        self.assertEqual(by_key["RN", 28, "NIGHT"]["EndTime"].date(), DATES[-1] + dt.timedelta(days=1))
        for name in RETAINED_ROLES:
            expected = sum(v * FIXTURE_HOURS for (r, _, _), v in pattern.items() if r == name)
            for lo, hi in ((1, 14), (15, 28)):
                self.assertAlmostEqual(sum(r["DemandHRS"] for r in rows if r["Role"] == name and lo <= r["Day"] <= hi), expected)

    def test_exclusions_do_not_change_retained_hours(self):
        a, p, c = fixture()
        original = pattern_from_publication(a, p, c)
        for name in ("AIN4", "CSM"):
            a.append(dict(a[0], Role=name, FTE=1000))
            p.append(dict(p[0], Role=name, RedistributedRosterFTE=1000))
        a.append(dict(a[0], Facility="JE", FTE=1000))
        p.append(dict(p[0], Facility="JE", RedistributedRosterFTE=1000))
        self.assertEqual(original, pattern_from_publication(a, p, c))
        for r in a + p:
            r["Role"] = {"RN": "Registered Nurse", "AIN": "Assistant in Nursing"}.get(r["Role"], r["Role"])
        self.assertEqual(original, pattern_from_publication(a, p, c))

    def test_en_replacement_preserves_rn_ain(self):
        a, p, c = fixture()
        baseline_pattern = pattern_from_publication(a, p, c, retained_roles=("RN", "AIN"))
        baseline = build_demand(baseline_pattern, DATES, SPANS, retained_roles=("RN", "AIN"))
        expected = build_demand(pattern_from_publication(a, p, c), DATES, SPANS)
        self.assertEqual([r for r in expected if r["Role"] in ("RN", "AIN")], baseline)
        self.assertEqual(sum(r["Role"] == "AINC4" for r in expected), 84)
        self.assertNotIn("EN", {r["Role"] for r in expected})
        for label in ("EN", "Enrolled Nurse", "AINC4", " en "):
            with self.subTest(label=label):
                pattern = pattern_from_publication(*fixture(third_role=label))
                self.assertEqual(build_demand(pattern, DATES, SPANS), expected)

    def test_third_role_requires_complete_unambiguous_data(self):
        a, p, c = fixture()
        with self.assertRaisesRegex(ValueError, "profile coverage"):
            pattern_from_publication([r for r in a if r["Role"] != "EN"],
                                     [r for r in p if r["Role"] != "EN"], c)
        p.append(dict(next(r for r in p if r["Role"] == "EN"), Role="AINC4"))
        with self.assertRaisesRegex(ValueError, "profile coverage"):
            pattern_from_publication(a, p, c)

    def test_invalid_publications_fail(self):
        for mutation in ("missing", "duplicate", "mismatch", "zero", "failed check", "nan", "wrong week"):
            with self.subTest(mutation=mutation):
                a, p, c = fixture()
                if mutation == "missing": p.pop()
                if mutation == "duplicate": a.append(copy.copy(a[0]))
                if mutation == "mismatch": a[0]["FTE"] += 1
                if mutation == "zero": a.pop(0)
                if mutation == "failed check": c[0]["Severity"] = "Error"
                if mutation == "nan": a[0]["FTE"] = float("nan")
                if mutation == "wrong week": a[0]["Week No"] += 1
                with self.assertRaises(ValueError):
                    pattern_from_publication(a, p, c)

    def test_invalid_calendar_and_duration_fail(self):
        pattern = pattern_from_publication(*fixture())
        for dates in (DATES[:14], DATES + [DATES[-1] + dt.timedelta(days=1)],
                      [d + dt.timedelta(days=1) for d in DATES], DATES[:5] + DATES[6:] + [DATES[-1]]):
            with self.assertRaises(ValueError):
                build_demand(pattern, dates, SPANS)
        spans = dict(SPANS)
        spans["RN", "AM"] = (dt.time(6), dt.time(14, 15), 7.6)
        with self.assertRaises(ValueError):
            build_demand(pattern, DATES, spans)

    def test_configured_duration_round_trip(self):
        baseline = build_demand(pattern_from_publication(*fixture()), DATES, SPANS)
        for hours in (7.6, 7.5, 8):
            shift_duration = standard_hours([{"ShiftDuration": hours}])
            a, p, c = fixture(shift_duration=shift_duration)
            # Same roster hours represented by a different standard-FTE unit.
            for row in a + p:
                for field in ("FTE", "HistoricalRosterFTE", "RedistributedRosterFTE"):
                    if field in row:
                        row[field] *= FIXTURE_HOURS / shift_duration
            pattern = pattern_from_publication(a, p, c, shift_duration=shift_duration)
            rows = build_demand(pattern, DATES, SPANS, shift_duration=shift_duration)
            for actual, expected in zip(rows, baseline):
                self.assertAlmostEqual(actual["DemandHRS"], expected["DemandHRS"])
                self.assertAlmostEqual(actual["DemandFTE"], expected["DemandFTE"])

    def test_invalid_settings_and_saved_basis(self):
        self.assertEqual(standard_hours([{"ShiftDuration": 1e308}]), 1e308)
        invalid = (None, True, "7.6", 0, -1, float("nan"), float("inf"), -float("inf"))
        for value in invalid:
            with self.subTest(value=value), self.assertRaises(ValueError):
                standard_hours([{"ShiftDuration": value}])
        for rows in ([], [{}], [{"ShiftDuration": 8}] * 2):
            with self.assertRaises(ValueError):
                standard_hours(rows)
        for collection in (0, 1):
            for value in (*invalid, FIXTURE_HOURS + 1, FIXTURE_HOURS * 60):
                publications = fixture()
                publications[collection][0]["ShiftDuration"] = value
                with self.assertRaisesRegex(ValueError, "saved FTE basis"):
                    pattern_from_publication(*publications)
            publications = fixture()
            del publications[collection][0]["ShiftDuration"]
            with self.assertRaisesRegex(ValueError, "saved FTE basis"):
                pattern_from_publication(*publications)

    def test_m_publication_contract(self):
        source = SOURCE.read_text(encoding="utf-8")
        names = re.findall(r'^shared (#[^=]+|\w+) =', source, re.M)
        self.assertEqual(len(names), len(set(names)))
        self.assertIn('shared #"Demand Roles" = {"RN", "AIN", "AINC4"};', source)
        self.assertNotIn('shared #"IMPORT LocRoleDayShift%"', source)
        published = source[source.index("shared ShiftUnitDemandHRS ="):]
        self.assertIn("DemandExtraction_CHECK", published)
        self.assertIn('error Error.Record("Demand Extraction validation failed"', published)
        for name in COLUMNS:
            self.assertIn('"' + name + '"', published)
        self.assertIn('Number.Mod(Duration.Days([Date] - Start), 14) + 1', source)
        self.assertIn('[SourceFTE] * ShiftDuration', source)
        self.assertIn('[DemandHRS] / [DurationOfShifts]', source)


def read_tables(path, wanted):
    """Read named-table cached values using ZIP/XML; never write a workbook."""
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    def parts(ref):
        match = re.fullmatch(r"([A-Z]+)([0-9]+)", ref)
        col = 0
        for ch in match[1]:
            col = col * 26 + ord(ch) - 64
        return col, int(match[2])
    with zipfile.ZipFile(path) as z:
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(si.itertext()) for si in ET.fromstring(z.read("xl/sharedStrings.xml"))]
        definitions = {}
        for name in z.namelist():
            if re.fullmatch(r"xl/tables/table[0-9]+.xml", name):
                root = ET.fromstring(z.read(name))
                if root.get("name") in wanted:
                    definitions[name] = root
        result = {}
        for rel_path in z.namelist():
            if not re.fullmatch(r"xl/worksheets/_rels/sheet[0-9]+.xml.rels", rel_path):
                continue
            for rel in ET.fromstring(z.read(rel_path)):
                target = rel.get("Target")
                target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl/worksheets", target))
                if target not in definitions:
                    continue
                table = definitions[target]
                lo, hi = map(parts, table.get("ref").split(":"))
                headers = [c.get("name") for c in table.find("s:tableColumns", ns)]
                sheet_path = rel_path.replace("/_rels", "").removesuffix(".rels")
                rows = {}
                for cell in ET.fromstring(z.read(sheet_path)).findall(".//s:sheetData/s:row/s:c", ns):
                    col, row = parts(cell.get("r"))
                    if not (lo[0] <= col <= hi[0] and lo[1] < row <= hi[1]):
                        continue
                    value = cell.find("s:v", ns)
                    if cell.get("t") == "inlineStr":
                        val = "".join(cell.find("s:is", ns).itertext())
                    elif value is None:
                        val = None
                    elif cell.get("t") == "s":
                        val = strings[int(value.text)]
                    elif cell.get("t") in ("str", "e"):
                        val = value.text
                    else:
                        val = float(value.text)
                    rows.setdefault(row, {})[headers[col - lo[0]]] = val
                result[table.get("name")] = [{h: data.get(h) for h in headers} for _, data in sorted(rows.items())]
        require(set(result) == set(wanted), "missing named tables")
        return result


def check_saved_workbooks():
    published = read_tables(UNIT / "1. Input/Demand-MasterRoster Manual Read.xlsx",
                           {"MinuteWorkersFTE_TABLE", "MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE", "MinuteWorkersFTE_CHECK"})
    schemas = []
    for name in ("MinuteWorkersFTE_TABLE", "MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE"):
        rows = published[name]
        columns = set(rows[0])
        schemas.append("DC" if {"DC Role", "DC Category"} <= columns else
                       "QFR" if "QFR Category" in columns else None)
        require(all(finite(r["Direct Care %"]) and 0 < r["Direct Care %"] <= 1
                    and finite(r["Week No"]) for r in rows), "saved source attributes")
    require(schemas[0] is not None and schemas[0] == schemas[1], "saved paired schema")
    settings = read_tables(UNIT / "2. Calculations/Settings Data.xlsx",
                           {"PermutationDimensions", "ShiftPeriod", "ShiftDuration"})
    shift_duration = standard_hours(settings["ShiftDuration"])
    pattern = pattern_from_publication(published["MinuteWorkersFTE_TABLE"],
                                       published["MinuteWorkersFTE_HISTORICAL_FORTNIGHT_TABLE"],
                                       published["MinuteWorkersFTE_CHECK"], shift_duration=shift_duration)
    calendar = [r for r in settings["PermutationDimensions"] if r["RolesList"] in RETAINED_ROLES]
    require(len(calendar) == 252, "saved calendar row count")
    require(len({(r["Date"], r["Shifts"], r["RolesList"]) for r in calendar}) == 252, "saved calendar duplicates")
    require(len({(r["Date"], r["Shifts"], r["Period"]) for r in calendar}) == 84, "saved period consistency")
    require({r["Period"] for r in calendar} == set(range(1, 85)), "saved periods")
    origin = dt.date(1899, 12, 30)
    dates = sorted({origin + dt.timedelta(days=r["Date"]) for r in calendar})
    require(all(r["Day"] == (origin + dt.timedelta(days=r["Date"]) - dates[0]).days + 1 for r in calendar), "saved Day")
    def time(value):
        seconds = round((value % 1) * 86400) % 86400
        return dt.time(seconds // 3600, seconds % 3600 // 60, seconds % 60)
    selected_spans = [r for r in settings["ShiftPeriod"] if r["Role"] in RETAINED_ROLES]
    require(len(selected_spans) == 9, "saved timing uniqueness")
    spans = {(r["Role"], r["ShiftPeriod"]): (time(r["StartDay"]), time(r["EndDay"]), r["DurationOfShifts"])
             for r in selected_spans}
    output = build_demand(pattern, dates, spans, shift_duration=shift_duration)
    # Verify actual saved IDs are preserved by the reference construction.
    saved_ids = {(origin + dt.timedelta(days=r["Date"]), r["Shifts"], r["RolesList"]): r["Period"] for r in calendar}
    require(all(r["Period"] == saved_ids[r["Date"], r["Shift"], r["Role"]] for r in output), "period preservation")
    print("Saved-table reference reconciliation: 252 rows; 28 days; 84 periods.")
    for name in RETAINED_ROLES:
        one = sum(r["DemandHRS"] for r in output if r["Role"] == name and r["Day"] <= 14)
        two = sum(r["DemandHRS"] for r in output if r["Role"] == name and r["Day"] > 14)
        require(abs(one - two) <= 1e-7, "fortnight repetition")
        print(f"BD {name}: {one:.9f} roster hours per fortnight; {one + two:.9f} over 28 days.")
    print("Independent reference calculation only; no M execution, synchronization or Excel refresh.")


if __name__ == "__main__":
    saved = "--saved-workbooks" in sys.argv
    result = unittest.main(argv=[sys.argv[0]], exit=False)
    if not result.result.wasSuccessful():
        sys.exit(1)
    if saved:
        check_saved_workbooks()
