"""Read saved XLSX table values without opening Excel or writing workbooks."""

import collections
import hashlib
import json
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ROOT = Path(__file__).resolve().parents[2]


def column_number(value):
    result = 0
    for char in value:
        result = result * 26 + ord(char) - 64
    return result


def bounds(ref):
    start, end = (ref.split(":") + [ref])[:2]
    a = re.fullmatch(r"([A-Z]+)(\d+)", start)
    b = re.fullmatch(r"([A-Z]+)(\d+)", end)
    return column_number(a[1]), int(a[2]), column_number(b[1]), int(b[2])


def xml_relationships(archive, part):
    folder, name = posixpath.split(part)
    relpath = posixpath.join(folder, "_rels", name + ".rels")
    if relpath not in archive.namelist():
        return {}
    result = {}
    for row in ET.fromstring(archive.read(relpath)):
        target = row.attrib["Target"]
        result[row.attrib["Id"]] = (
            target.lstrip("/") if target.startswith("/") else
            posixpath.normpath(posixpath.join(folder, target))
        )
    return result


def read_book(relative, selected=None):
    path = ROOT / relative
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    result = {"path": relative, "sha256": before, "tables": []}
    with ZipFile(path) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            for item in ET.fromstring(archive.read("xl/sharedStrings.xml")):
                shared.append("".join(item.itertext()))
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = xml_relationships(archive, "xl/workbook.xml")
        for sheet in workbook.find("s:sheets", NS):
            sheet_path = relationships[sheet.attrib[f"{{{REL}}}id"]]
            sheet_xml = ET.fromstring(archive.read(sheet_path))
            sheet_relationships = xml_relationships(archive, sheet_path)
            parts = sheet_xml.find("s:tableParts", NS)
            if parts is None:
                continue
            for part in parts:
                table_path = sheet_relationships[part.attrib[f"{{{REL}}}id"]]
                table = ET.fromstring(archive.read(table_path))
                name = table.attrib["displayName"]
                cols = [col.attrib["name"] for col in table.find("s:tableColumns", NS)]
                left, top, right, bottom = bounds(table.attrib["ref"])
                info = {"name": name, "sheet": sheet.attrib["name"],
                        "ref": table.attrib["ref"], "columns": cols,
                        "rows": max(0, bottom - top - int(table.attrib.get("totalsRowCount", "0")))}
                if selected is not None and (selected == "*" or name in selected):
                    rows = []
                    for row in sheet_xml.findall("s:sheetData/s:row", NS):
                        index = int(row.attrib["r"])
                        if not top < index <= bottom:
                            continue
                        values = {}
                        for cell in row:
                            match = re.fullmatch(r"([A-Z]+)(\d+)", cell.attrib["r"])
                            column = column_number(match[1])
                            if not left <= column <= right:
                                continue
                            value = cell.find("s:v", NS)
                            kind = cell.attrib.get("t")
                            if kind == "inlineStr":
                                value = "".join(cell.find("s:is", NS).itertext())
                            elif value is None:
                                value = None
                            elif kind == "s":
                                value = shared[int(value.text)]
                            elif kind == "b":
                                value = value.text == "1"
                            elif kind in {"str", "e"}:
                                value = value.text
                            else:
                                raw = value.text
                                value = float(raw) if raw and any(x in raw for x in ".eE") else int(raw) if raw else None
                            values[cols[column-left]] = value
                        rows.append({col: values.get(col) for col in cols})
                    info["data"] = rows
                result["tables"].append(info)
    result["unchanged"] = before == hashlib.sha256(path.read_bytes()).hexdigest()
    return result


def main():
    prefix = "CLIENT/DATExx-Whiddon/UNITS/"
    files = [("JH-RY", "2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx"),
             ("JH-RY", "2. Calculations/Capacity-ShiftAvailability.xlsx"),
             ("JH-RY", "2. Calculations/StaffListMaster.xlsx"),
             ("JH-RY", "2. Calculations/Settings Data.xlsx"),
             ("BD", "2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx"),
             ("TE", "2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx")]
    if "--diagnose" not in sys.argv:
        for unit, suffix in files:
            print(json.dumps(read_book(prefix + unit + "/" + suffix), ensure_ascii=False))
        return
    selected = {"CapacityDiagnostics_SUMMARY", "CapacityDiagnostics_DETAILS", "IMPORT_ResDayShift",
                "ResPeriodWDTABLE", "ResPeriodAvailabilityTABLE", "C__ResSingleShiftDayTABLE",
                "ResourcePeriodTABLE_empty", "Table_Masterlist", "ResDayShift", "PermutationDimensions",
                "ResDayShift_CHECK", "ResourceContract_CHECK", "RolePathTABLE"}
    books = {unit + "/" + suffix: read_book(prefix + unit + "/" + suffix, selected)
             for unit, suffix in files}
    def tables(book):
        return {table["name"]: table for table in book["tables"] if "data" in table}
    def data(table):
        return [row for row in table["data"] if any(value is not None for value in row.values())]
    def counted(rows, columns):
        return [{"key": key, "count": count} for key, count in
                collections.Counter(tuple(row.get(col) for col in columns) for row in rows).most_common()]
    def dates(rows):
        values = [row["Date"] for row in rows if isinstance(row.get("Date"), (int, float))]
        return {"min_serial": min(values) if values else None, "max_serial": max(values) if values else None}
    for key, book in books.items():
        found = tables(book)
        summary = {"book": key, "unchanged": book["unchanged"]}
        for name, table in found.items():
            rows = data(table)
            item = {"rows": len(rows), "sheet": table["sheet"], "range": table["ref"]}
            if name.endswith("SUMMARY") or name.endswith("CHECK"):
                item["failures"] = [row for row in rows if row.get("Status") != "Pass"]
            elif name.endswith("DETAILS"):
                item["reasons"] = counted(rows, ["Check", "Reason", "Action"])
                item["first_per_reason"] = list({row.get("Reason"): {"excel_row": index + 2, **{k:v for k,v in row.items() if k != "Name"}}
                                                 for index, row in reversed(list(enumerate(rows)))}.values())[:15]
            elif name in {"ResDayShift", "IMPORT_ResDayShift", "Table_Masterlist", "PermutationDimensions"}:
                item["roles"] = counted(rows, ["Role"] if name != "PermutationDimensions" else ["RolesList"])
                item["facilities"] = counted(rows, ["Facility-Abbrev"])
                item["dates"] = dates(rows)
            elif name == "RolePathTABLE":
                item["values"] = rows
            elif name == "ResourcePeriodTABLE_empty":
                item["spare_reasons"] = counted(rows, ["SpareIssueReason"])
                item["original_availability"] = counted(rows, ["OriginalAvailability"])
            else:
                item["availability"] = counted(rows, ["Availability"])
            summary[name] = item
        print(json.dumps(summary, ensure_ascii=False))
    ain = tables(books["JH-RY/2. Calculations/AIN/CapacityDistrib(A.1)-shifts.xlsx"])
    staff = tables(books["JH-RY/2. Calculations/StaffListMaster.xlsx"])
    settings = tables(books["JH-RY/2. Calculations/Settings Data.xlsx"])
    workers = [row for row in data(staff["Table_Masterlist"]) if row["Role"] == "AIN"]
    names = collections.Counter((row["Name"], row["Role"]) for row in workers)
    periods = {(row["Date"], row["Shifts"], row["RolesList"]) for row in data(settings["PermutationDimensions"])}
    availability = data(ain["IMPORT_ResDayShift"])
    identity_matches = [row for row in availability if names[(row["Name"], row["Role"])] == 1]
    period_matches = [row for row in identity_matches if (row["Date"], row["Shift"], row["Role"]) in periods]
    print(json.dumps({"join_reproduction": {"AIN_master_workers": len(workers),
          "saved_A1_availability_rows": len(availability), "unique_name_role_matches": len(identity_matches),
          "date_shift_role_matches": len(period_matches),
          "unmatched_names": len({row["Name"] for row in availability if names[(row["Name"], row["Role"])] != 1})}}))


if __name__ == "__main__":
    main()
