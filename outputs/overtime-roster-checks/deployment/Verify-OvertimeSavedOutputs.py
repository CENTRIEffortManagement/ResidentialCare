"""Read-only saved-output acceptance for the three authorised overtime workbooks.

Usage (from the repository root, after all three refreshes succeeded):
    python outputs/overtime-roster-checks/deployment/Verify-OvertimeSavedOutputs.py \
        outputs/overtime-roster-checks/deployment/<run-folder>

Only standard-library ZIP/XML reads are used. Workbook bytes are never written.
The JSON contains aggregate counts and technical contracts, never employee rows.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile


REPO = Path(__file__).resolve().parents[3]
TARGETS = {
    "Core": "CLIENT/DATExx-Whiddon/2. Calculations/OvertimeRosterChecks/OvertimeRosterChecks.xlsx",
    "Shifts": "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/Shifts.xlsx",
    "Allocation": "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/AllocationByShiftAverage.xlsx",
}
SOURCES = {
    "Core": "Workflows/ResidentialCare/OvertimeRosterChecks/OvertimeRosterChecks.pq",
    "Shifts": "Workflows/ResidentialCare/OvertimeRosterChecks/BD/Shifts.pq",
}
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
Q = lambda name: "{" + MAIN + "}" + name
EPS = Decimal("0.0000001")
DAY_EPS = Decimal("0.00000001")
DATE_COLUMNS = {
    "DateTime", "TimeDate", "DayDate", "ShiftDate", "SourceDate", "ShiftStart", "ShiftFinish",
    "PreviousShiftFinish", "PayrollFortnightStart", "PayrollFortnightEnd", "SegmentStart", "SegmentFinish",
}
MATCH_FIELDS = ("SourceUnit", "AllocationCode", "Role", "ShiftStart", "ShiftFinish")
CHECK_FIELDS = (
    "SourceShiftID", "DailyStatus", "DailyExcessHours", "DailyHoursFlag", "GapStatus", "GapHours",
    "MinimumBreakFlag", "MinBreakStatus", "ShortGapReviewFlag", "ShortGapStatus", "DaysOffStatus",
    "DaysOffFlag", "FortnightStatus", "FortnightExcessHours", "OvertimeHours", "OvertimeStatus",
    "AnyRuleFlag", "CoverageStatus", "TriggerSummary", "ResultReasons",
)
REMOVED_FIELDS = {
    "PreviousRosterShiftPeriod", "NightTransitionFlag", "NightDutyFreeDays", "NightDaysFlag",
    "NightGapStatus", "NightStatus", "NightTransitionStatus",
}


class NotReady(Exception):
    pass


@dataclass(frozen=True)
class CellError:
    code: str


def repo_path(relative: str) -> Path:
    candidate = Path(relative)
    if candidate.is_absolute():
        raise ValueError("Paths must be repo-relative.")
    resolved = (REPO / candidate).resolve()
    if not resolved.is_relative_to(REPO):
        raise ValueError("Path leaves the repository.")
    return resolved


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def normal(value):
    if isinstance(value, CellError):
        return ["ExcelError", value.code]
    if isinstance(value, Decimal):
        return ["Number", str(value.quantize(Decimal("0.0000000001")).normalize())]
    if isinstance(value, bool):
        return ["Boolean", value]
    return value


def fingerprint(values) -> bytes:
    payload = json.dumps([normal(value) for value in values], ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).digest()


def number(value):
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return Decimal(int(value))
    return None


def close(left, right, tolerance=EPS):
    left, right = number(left), number(right)
    return left is not None and right is not None and abs(left - right) <= tolerance


def is_true(value):
    return value is True or value == Decimal(1)


def iso_serial(value):
    if not isinstance(value, Decimal):
        return None
    return (datetime(1899, 12, 30) + timedelta(seconds=float(value * 86400))).isoformat(timespec="seconds")


def json_default(value):
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    raise TypeError("Unsupported aggregate report type.")


def relationships(archive: ZipFile, name: str) -> dict:
    if name not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(name))
    return {item.attrib["Id"]: item.attrib for item in root.findall("{" + REL + "}Relationship")}


def target_part(source_part: str, target: str) -> str:
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def column_index(letters: str) -> int:
    result = 0
    for letter in letters:
        result = result * 26 + ord(letter) - ord("A") + 1
    return result


def cell_position(reference: str):
    match = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", reference)
    if match is None:
        raise ValueError("Unsupported worksheet cell reference.")
    return column_index(match[1]), int(match[2])


class SavedBook:
    """Keep table metadata small and stream worksheet rows; shared strings load once per book."""

    def __init__(self, path: Path):
        self.path = path
        self.archive = ZipFile(path, "r")
        self.strings = None
        root = ET.fromstring(self.archive.read("xl/workbook.xml"))
        props = root.find(Q("workbookPr"))
        self.date1904 = props is not None and props.attrib.get("date1904", "0").lower() in {"1", "true"}
        book_rels = relationships(self.archive, "xl/_rels/workbook.xml.rels")
        self.tables = {}
        for sheet in root.findall(Q("sheets") + "/" + Q("sheet")):
            relation = book_rels[sheet.attrib["{" + DOC_REL + "}id"]]
            if relation.get("TargetMode") == "External":
                continue
            part = target_part("xl/workbook.xml", relation["Target"])
            rel_part = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
            for item in relationships(self.archive, rel_part).values():
                if not item["Type"].endswith("/table") or item.get("TargetMode") == "External":
                    continue
                definition = ET.fromstring(self.archive.read(target_part(part, item["Target"])))
                name = definition.attrib["name"]
                if name.casefold() in self.tables:
                    raise ValueError("Duplicate saved Excel table names.")
                start, end = definition.attrib["ref"].split(":") if ":" in definition.attrib["ref"] else (definition.attrib["ref"],) * 2
                left, top = cell_position(start)
                right, bottom = cell_position(end)
                headers = [item.attrib["name"] for item in definition.findall(Q("tableColumns") + "/" + Q("tableColumn"))]
                self.tables[name.casefold()] = {
                    "name": name, "sheet": sheet.attrib["name"], "part": part, "headers": headers,
                    "left": left, "right": right, "top": top + int(definition.attrib.get("headerRowCount", "1")),
                    "bottom": bottom - int(definition.attrib.get("totalsRowCount", "0")),
                }

    def close(self):
        self.strings = None
        self.archive.close()

    def shared_strings(self):
        if self.strings is not None:
            return self.strings
        self.strings = []
        if "xl/sharedStrings.xml" in self.archive.namelist():
            with self.archive.open("xl/sharedStrings.xml") as stream:
                stack = []
                for event, element in ET.iterparse(stream, events=("start", "end")):
                    if event == "start":
                        stack.append(element)
                    else:
                        if element.tag == Q("si"):
                            self.strings.append("".join(text.text or "" for child in element if child.tag != Q("rPh") for text in child.iter(Q("t"))))
                            if len(stack) > 1:
                                stack[-2].remove(element)
                            element.clear()
                        stack.pop()
        return self.strings

    def value(self, cell, header):
        kind = cell.attrib.get("t", "n")
        raw = cell.findtext(Q("v"))
        if kind == "e":
            return CellError(raw or "#ERROR")
        if kind == "inlineStr":
            return "".join(text.text or "" for text in cell.iter(Q("t")))
        if raw is None:
            return None
        if kind == "s":
            try:
                return self.shared_strings()[int(raw)]
            except (ValueError, IndexError):
                return CellError("#INVALID_SHARED_STRING")
        if kind == "b":
            return raw == "1"
        if kind == "d":
            try:
                stamp = datetime.fromisoformat(raw.replace("Z", "+00:00"))
                if stamp.tzinfo is not None:
                    stamp = stamp.astimezone(timezone.utc).replace(tzinfo=None)
                return Decimal(str((stamp - datetime(1899, 12, 30)).total_seconds())) / 86400
            except ValueError:
                return CellError("#INVALID_DATE_CACHE")
        if kind in {"str", "inlineStr"}:
            return raw
        try:
            result = Decimal(raw)
            if not result.is_finite():
                return CellError("#NONFINITE_NUMBER")
            return result + 1462 if self.date1904 and header in DATE_COLUMNS else result
        except InvalidOperation:
            return CellError("#INVALID_NUMBER_CACHE")

    def rows(self, table):
        self.shared_strings()
        columns = {table["left"] + index: name for index, name in enumerate(table["headers"])}
        with self.archive.open(table["part"]) as stream:
            stack = []
            for event, element in ET.iterparse(stream, events=("start", "end")):
                if event == "start":
                    stack.append(element)
                    continue
                if element.tag == Q("row"):
                    row_number = int(element.attrib["r"])
                    if table["top"] <= row_number <= table["bottom"]:
                        values = {}
                        for cell in element.findall(Q("c")):
                            index, _ = cell_position(cell.attrib["r"])
                            if index in columns:
                                header = columns[index]
                                values[header] = self.value(cell, header)
                        # Excel may persist one completely blank body row for an empty query table.
                        if any(value is not None and value != "" for value in values.values()):
                            yield values
                    if len(stack) > 1:
                        stack[-2].remove(element)
                    element.clear()
                    if row_number > table["bottom"]:
                        break
                stack.pop()


class Verification:
    def __init__(self):
        self.checks = []
        self.tables = {}
        self.warnings = []

    def check(self, name, passed, **counts):
        self.checks.append({"Check": name, "Passed": bool(passed), **counts})

    def table(self, book, name, required):
        table = book.tables.get(name.casefold())
        missing = sorted(set(required) - set(table["headers"] if table else ()))
        self.check("Loaded table contract: " + name, table is not None and not missing, MissingColumns=missing)
        if table is not None:
            self.check("Removed pre-night fields absent: " + name, not (REMOVED_FIELDS & set(table["headers"])))
        return table

    def scan(self, book, table):
        stats = {"Rows": 0, "ExcelErrorCells": 0, "ExcelErrorsByColumn": defaultdict(Counter), "StatusCounts": defaultdict(Counter)}
        self.tables[table["name"]] = stats
        status_columns = [name for name in table["headers"] if name.endswith("Status") or name.endswith("Flag")]
        date_bounds = {}
        for row in book.rows(table):
            stats["Rows"] += 1
            for name, value in row.items():
                if isinstance(value, CellError):
                    stats["ExcelErrorCells"] += 1
                    stats["ExcelErrorsByColumn"][name][value.code] += 1
                elif name in DATE_COLUMNS and isinstance(value, Decimal):
                    bounds = date_bounds.setdefault(name, [value, value])
                    bounds[0], bounds[1] = min(bounds[0], value), max(bounds[1], value)
            for name in status_columns:
                value = row.get(name)
                label = "(Excel error)" if isinstance(value, CellError) else "(blank)" if value is None or value == "" else str(value)
                stats["StatusCounts"][name][label] += 1
            if row.get("ResultReasons"):
                stats["RowsWithResultReasons"] = stats.get("RowsWithResultReasons", 0) + 1
            if "BeforeNight" in str(row.get("TriggerSummary") or ""):
                stats["RemovedTriggerRows"] = stats.get("RemovedTriggerRows", 0) + 1
            yield row
        stats["SavedDateRanges"] = {name: {"Min": iso_serial(bounds[0]), "Max": iso_serial(bounds[1])} for name, bounds in date_bounds.items()}
        self.check("No Excel error cells: " + table["name"], stats["ExcelErrorCells"] == 0, ErrorCells=stats["ExcelErrorCells"])
        self.check("No removed pre-night trigger: " + table["name"], stats.get("RemovedTriggerRows", 0) == 0)


def ready_receipts(run: Path):
    receipts = {}
    for target, expected in TARGETS.items():
        path = run / (target + "-refresh-receipt.json")
        if not path.is_file():
            raise NotReady("All three successful refresh receipts are required before workbook inspection.")
        receipt = load_json(path)
        if receipt.get("Status") != "Passed" or receipt.get("ExitCode") != 0 or receipt.get("Saved") is not True or receipt.get("Closed") is not True:
            raise NotReady("All three targets must have Passed, Saved and Closed refresh receipts.")
        if receipt.get("Workbook", "").replace("\\", "/").casefold() != expected.casefold():
            raise NotReady("A refresh receipt names an alternative workbook.")
        receipts[target] = receipt
    for producer, consumer in (("Core", "Shifts"), ("Shifts", "Allocation")):
        finish = datetime.fromisoformat(receipts[producer]["FinishedAtUTC"].replace("Z", "+00:00"))
        start = datetime.fromisoformat(receipts[consumer]["StartedAtUTC"].replace("Z", "+00:00"))
        if finish > start:
            raise NotReady("Refresh receipts do not establish dependency order.")
    for target in ("Core", "Shifts"):
        sync = load_json(run / (target + "-sync-receipt.json"))
        if sync.get("Source", "").replace("\\", "/").casefold() != SOURCES[target].casefold() or sync.get("Workbook", "").replace("\\", "/").casefold() != TARGETS[target].casefold():
            raise NotReady("A synchronization receipt names an alternative source or workbook.")
        if sync.get("ReExtractionMatches") is not True or sync.get("Saved") is not True or sync.get("Closed") is not True:
            raise NotReady("Verified source synchronization receipts are required.")
        if sha256(repo_path(sync["Source"])) != sync["SourceSHA256"].upper():
            raise NotReady("Approved M source changed after synchronization.")
        receipts[target]["VerifiedSource"] = {"Path": SOURCES[target], "SHA256": sync["SourceSHA256"], "ReExtractionMatches": True}
    return receipts


def verify(receipts, report):
    v = Verification()
    report.update({"Checks": v.checks, "Tables": v.tables, "Warnings": v.warnings})
    report["Phase"] = "Verifying saved workbook hashes"
    hashes = {}
    for target, relative in TARGETS.items():
        hashes[target] = sha256(repo_path(relative))
        if hashes[target] != receipts[target]["WorkbookSHA256After"].upper():
            raise NotReady("A saved workbook differs from its completed refresh receipt.")
    report["Workbooks"] = {target: {"Path": relative, "SHA256": hashes[target], **({"VerifiedSource": receipts[target]["VerifiedSource"]} if "VerifiedSource" in receipts[target] else {})} for target, relative in TARGETS.items()}
    report["Phase"] = "Reading Core output tables"
    core = SavedBook(repo_path(TARGETS["Core"]))
    shifts = allocation = None
    try:
        core_table = v.table(core, "OvertimeRosterChecks_ShiftResults", set(MATCH_FIELDS) | set(CHECK_FIELDS) | {"EmployeeKey", "QualifyingNetHours", "NetRosterHours", "RosterShiftPeriod"})
        fn_table = v.table(core, "OvertimeRosterChecks_FortnightResults", {"EmployeeKey", "PayrollFortnightStart", "PayrollFortnightEnd", "NetRosterHours", "FortnightExcessHours", "DutyFreeDays", "DaysShort", "CoverageStatus", "FortnightStatus", "DaysOffStatus"})
        segment_table = v.table(core, "OvertimeRosterChecks_ExcessSegments", {"SourceShiftID", "NetWorkStartHours", "NetWorkEndHours", "SegmentHours", "SegmentStart", "SegmentFinish", "DailyTrigger", "FortnightTrigger", "TimingStatus", "OvertimeStatus", "TriggerSummary"})
        core_ids, core_matches = {}, defaultdict(list)
        duplicate_ids = missing_ids = 0
        if core_table:
            for row in v.scan(core, core_table):
                shift_id = row.get("SourceShiftID")
                if not shift_id or isinstance(shift_id, CellError):
                    missing_ids += 1
                    continue
                duplicate_ids += shift_id in core_ids
                core_ids[shift_id] = {name: row.get(name) for name in set(MATCH_FIELDS) | set(CHECK_FIELDS) | {"QualifyingNetHours"}}
                core_matches[fingerprint(row.get(name) for name in MATCH_FIELDS)].append(shift_id)
            v.check("Core SourceShiftID is present and unique", missing_ids == duplicate_ids == 0, MissingIDs=missing_ids, DuplicateIDs=duplicate_ids, UniquePhysicalShifts=len(core_ids))
            if not core_ids:
                v.warnings.append("Core shift output is empty; no business population was established.")
        if fn_table:
            keys, duplicate_fn = set(), 0
            for row in v.scan(core, fn_table):
                key = fingerprint((row.get("EmployeeKey"), row.get("PayrollFortnightStart")))
                duplicate_fn += key in keys
                keys.add(key)
            v.check("Employee-fortnight publication keys are unique", duplicate_fn == 0, DuplicateKeys=duplicate_fn, EmployeeFortnights=len(keys))
        ranges, segment_hours = defaultdict(list), defaultdict(Decimal)
        invalid_segments = unknown_clock = 0
        if segment_table:
            for row in v.scan(core, segment_table):
                shift_id = row.get("SourceShiftID")
                left, right, hours = (number(row.get(name)) for name in ("NetWorkStartHours", "NetWorkEndHours", "SegmentHours"))
                valid = shift_id in core_ids and left is not None and right is not None and hours is not None and left >= 0 and right > left and hours > 0 and close(right - left, hours)
                valid = valid and (is_true(row.get("DailyTrigger")) or is_true(row.get("FortnightTrigger")))
                maximum = number(core_ids.get(shift_id, {}).get("QualifyingNetHours"))
                if maximum is not None and right is not None:
                    valid = valid and right <= maximum + EPS
                if row.get("TimingStatus") == "Known":
                    start, finish = number(row.get("SegmentStart")), number(row.get("SegmentFinish"))
                    valid = valid and start is not None and finish is not None and finish > start and close((finish - start) * 24, hours, Decimal("0.000001"))
                else:
                    unknown_clock += 1
                invalid_segments += not valid
                if left is not None and right is not None and hours is not None:
                    ranges[shift_id].append((left, right))
                    segment_hours[shift_id] += hours
            overlap_count = 0
            union_hours = Decimal(0)
            for values in ranges.values():
                previous = None
                for left, right in sorted(values):
                    overlap_count += previous is not None and left < previous - EPS
                    union_hours += max(Decimal(0), right - max(left, previous if previous is not None else left))
                    previous = right if previous is None else max(previous, right)
            v.check("Excess ranges are positive, bounded and correctly timed when known", invalid_segments == 0, InvalidSegments=invalid_segments, UnknownClockSegments=unknown_clock)
            v.check("Excess net ranges are disjoint and counted once", overlap_count == 0 and close(union_hours, sum(segment_hours.values(), Decimal(0))), OverlappingRanges=overlap_count, ObservedExcessNetHours=sum(segment_hours.values(), Decimal(0)), UnionExcessNetHours=union_hours)
            mismatches = sum(1 for shift_id, row in core_ids.items() if number(row.get("OvertimeHours")) is not None and not close(row["OvertimeHours"], segment_hours[shift_id]))
            v.check("Known shift overtime totals equal excess interval union", mismatches == 0, MismatchedShifts=mismatches)
        core.close()
        core = None

        report["Phase"] = "Reading Shifts plotting tables"
        shifts = SavedBook(repo_path(TARGETS["Shifts"]))
        allocation = SavedBook(repo_path(TARGETS["Allocation"]))
        plot = v.table(shifts, "OvertimeRosterChecks", set(MATCH_FIELDS) | set(CHECK_FIELDS) | {"Name", "DateTime", "StaffCount", "ShiftType", "LocalShiftKey", "OvertimeMatchCount", "OvertimeMatchStatus", "RosterShiftPeriod"})
        names = v.table(shifts, "NAMES_INTERVALSLIST", set(CHECK_FIELDS) | {"LocalShiftKey", "AllocatedIntervals", "Attribute", "Role1", "TimeDate", "RealDuration", "IntervalDurationTemp"})
        ria = v.table(allocation, "ResourceShiftAllocationY", set(CHECK_FIELDS) | {"LocalShiftKey", "AllocatedIntervals", "Attribute", "Role1", "TimeDate", "RealDuration", "IntervalDuration", "ResEffectiveIntervalEffort"})
        intervals = v.table(shifts, "IntervalsList", {"IntervalID", "Role", "TimeDate", "Attribute"})
        if intervals:
            for _ in v.scan(shifts, intervals):
                pass
        local = {}
        attr_conflicts = missing_local = match_errors = 0
        match_counts = Counter()
        if plot:
            attribute_columns = sorted(set(plot["headers"]) - {"DateTime", "StaffCount"})
            inherited_columns = sorted(set(CHECK_FIELDS) & set(plot["headers"]))
            for row in v.scan(shifts, plot):
                key = row.get("LocalShiftKey")
                if not key or isinstance(key, CellError):
                    missing_local += 1
                    continue
                attributes = fingerprint(row.get(name) for name in attribute_columns)
                group = local.setdefault(key, {"rows": 0, "attributes": attributes, "points": [], "start": row.get("ShiftStart"), "finish": row.get("ShiftFinish"), "checks": fingerprint(row.get(name) for name in inherited_columns), "SourceShiftID": row.get("SourceShiftID")})
                group["rows"] += 1
                attr_conflicts += attributes != group["attributes"]
                if len(group["points"]) < 5:
                    group["points"].append((row.get("DateTime"), row.get("StaffCount")))
                if group["rows"] == 1:
                    status = row.get("OvertimeMatchStatus")
                    match_counts[str(status)] += 1
                    candidates = [] if row.get("AllocationCode") in {None, ""} or row.get("Role") is None else core_matches.get(fingerprint(row.get(name) for name in MATCH_FIELDS), [])
                    expected = "Matched" if len(candidates) == 1 else "Unmatched" if not candidates else "Ambiguous"
                    valid = status == expected and close(row.get("OvertimeMatchCount"), Decimal(len(candidates)))
                    if len(candidates) == 1:
                        valid = valid and row.get("SourceShiftID") == candidates[0]
                        valid = valid and fingerprint(row.get(name) for name in inherited_columns) == fingerprint(core_ids[candidates[0]].get(name) for name in inherited_columns)
                    else:
                        valid = valid and row.get("SourceShiftID") in {None, ""}
                    match_errors += not valid
            bad_points = 0
            for group in local.values():
                start, finish = number(group["start"]), number(group["finish"])
                expected = [] if start is None or finish is None else [(start, 0), (start + Decimal(1) / 1440, 1), (finish - Decimal(1) / 1440, 1), (finish, 0)]
                valid = group["rows"] == 4
                for stamp, count in group["points"]:
                    position = next((index for index, pair in enumerate(expected) if close(stamp, pair[0], DAY_EPS) and close(count, Decimal(pair[1]))), None)
                    if position is None:
                        valid = False
                    else:
                        expected.pop(position)
                bad_points += not valid or bool(expected)
            v.check("Every local physical row has exactly four conserved plotting points", missing_local == bad_points == attr_conflicts == 0, LocalPhysicalShifts=len(local), MissingLocalKeys=missing_local, InvalidFourPointGroups=bad_points, RepeatedAttributeConflicts=attr_conflicts, PlottingRows=v.tables[plot["name"]]["Rows"])
            v.check("Physical five-field core join is truthful without fan-out", match_errors == 0, InvalidJoinGroups=match_errors, MatchCounts=dict(match_counts))

        inherited_mismatch = missing_interval_key = 0
        report["Phase"] = "Comparing saved interval inheritance"
        expected_rows, observed_rows = Counter(), Counter()
        expected_count = observed_count = 0
        if names and ria:
            inherited = sorted((set(names["headers"]) & set(ria["headers"])) | {"IntervalDuration"})
            check_columns = sorted(set(CHECK_FIELDS) & set(names["headers"]))
            for row in v.scan(shifts, names):
                key = row.get("LocalShiftKey")
                missing_interval_key += key not in local
                if key in local:
                    inherited_mismatch += fingerprint(row.get(name) for name in check_columns) != local[key]["checks"]
                # This is the existing M filter, including M's null <> 0 behaviour.
                if row.get("RealDuration") != Decimal(0):
                    row["IntervalDuration"] = row.get("IntervalDurationTemp")
                    expected_rows[fingerprint(row.get(name) for name in inherited)] += 1
                    expected_count += 1
            shifts.close()
            shifts = None
            for row in v.scan(allocation, ria):
                observed_rows[fingerprint(row.get(name) for name in inherited)] += 1
                observed_count += 1
            missing_rows = sum((expected_rows - observed_rows).values())
            additional_rows = sum((observed_rows - expected_rows).values())
            v.check("Interval rows retain their own local physical identity and check attributes", missing_interval_key == inherited_mismatch == 0, MissingLocalKeyRows=missing_interval_key, AttributeMismatchRows=inherited_mismatch)
            v.check("Allocation interval publication conserves filtered Shifts rows and inherited values", expected_rows == observed_rows, ExpectedFilteredRows=expected_count, ActualAllocationRows=observed_count, MissingRows=missing_rows, AdditionalRows=additional_rows, ComparedColumns=inherited)
        for target, relative in TARGETS.items():
            v.check("Workbook unchanged during read-only verification: " + target, sha256(repo_path(relative)) == hashes[target])
        report.update({"Status": "Passed" if all(check["Passed"] for check in v.checks) else "Failed", "Checks": v.checks, "Tables": v.tables, "Warnings": v.warnings})
        report["Phase"] = "Complete"
    finally:
        for book in (core, shifts, allocation):
            if book is not None:
                book.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("run_folder", help="Repo-relative completed deployment evidence folder.")
    parser.add_argument("--output-json", help="Repo-relative JSON destination inside that evidence folder.")
    args = parser.parse_args()
    run = repo_path(args.run_folder)
    output = repo_path(args.output_json) if args.output_json else run / "saved-output-verification.json"
    if output.suffix.lower() != ".json" or not output.is_relative_to(run):
        parser.error("Output must be a JSON file inside the selected deployment evidence folder.")
    report = {"Status": "NotReady", "ReadOnly": True, "VerifiedAtUTC": datetime.now(timezone.utc).isoformat(), "RunFolder": run.relative_to(REPO).as_posix(), "Scope": "Saved tables in the three authorised targets only; no Excel process, query execution or workbook mutation.", "Limitations": ["Reads saved cached table values, not unsaved Excel state or live SharePoint inputs.", "Business Unknown, unmatched and ambiguous results are reported separately from technical cell errors.", "Does not establish payroll entitlement or verify unrefreshed upstream input currency.", "Settings anchor derivation is established by the approved M source; only published fortnight buckets are inspected here.", "Four-point conservation is checked per published LocalShiftKey; no additional upstream workbook is inspected to recount input shifts."]}
    try:
        receipts = ready_receipts(run)  # No workbook is opened before this succeeds.
        verify(receipts, report)
    except NotReady as exc:
        report["Status"] = "NotReady"
        report["Reason"] = str(exc)
    except Exception as exc:
        report["Status"] = "Failed"
        # Do not serialize raw exception values: a malformed employee cell may be involved.
        report["FailureType"] = type(exc).__name__
        report["Reason"] = "Saved-output verification could not complete; no workbook was changed."
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, indent=2, default=json_default) + "\n", encoding="utf-8")
    print(json.dumps({"Status": report["Status"], "Output": output.relative_to(REPO).as_posix(), "Checks": len(report.get("Checks", [])), "FailedChecks": sum(not item["Passed"] for item in report.get("Checks", []))}, indent=2))
    return 0 if report["Status"] == "Passed" else 2 if report["Status"] == "NotReady" else 1


if __name__ == "__main__":
    sys.exit(main())
