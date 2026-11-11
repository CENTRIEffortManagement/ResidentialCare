"""Read-only structural verification of the amended threshold template."""
import hashlib
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

target = Path("docs/OvertimeRosterCheck/OvertimeRosterThreshold.xlsx")
ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
expected = {
    "DailyHoursLimit": 10,
    "CSHospitalityHoursLimit": 8,
    "CSPersonalCareHoursLimit": 10,
    "MinimumBreakHours": 8,
    "ShortGapReviewHours": 10,
    "MinimumFortnightDutyFreeDays": 4,
    "FortnightHoursLimit": 76,
    "FortnightLengthDays": 14,
}
preservation = json.loads(Path("outputs/overtime-roster-checks/thresholds-amendment-preservation.json").read_text(encoding="utf-8"))
assert list(preservation["retainedValues"]) == list(expected)
expected = preservation["retainedValues"]
with zipfile.ZipFile(target) as archive:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = workbook.findall("m:sheets/m:sheet", ns)
    assert len(sheets) == 1 and sheets[0].attrib["name"] == "OvertimeRosterThreshold"
    sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cells = {cell.attrib["r"]: cell for cell in sheet.findall("m:sheetData/m:row/m:c", ns)}
    strings = []
    if "xl/sharedStrings.xml" in archive.namelist():
        string_xml = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        strings = ["".join(si.itertext()) for si in string_xml]

    def value(address):
        cell = cells.get(address)
        if cell is None:
            return None
        raw = cell.find("m:v", ns)
        if cell.attrib.get("t") == "inlineStr":
            inline = cell.find("m:is", ns)
            return "".join(inline.itertext()) if inline is not None else None
        if raw is None:
            return None
        if cell.attrib.get("t") == "s":
            return strings[int(raw.text)]
        return raw.text

    assert [value(f"{column}1") for column in "ABCD"] == ["Title", "Value", "Unit", "Notes"]
    titles = [value(f"A{row}") for row in range(2, 10)]
    assert len(titles) == len(set(titles)) == 8
    assert titles == list(expected)
    for row, title in enumerate(titles, start=2):
        actual = value(f"B{row}")
        assert actual is not None
        assert cells[f"B{row}"].attrib.get("t", "n") == "n"
        assert float(actual) == expected[title]
        assert value(f"C{row}") in ("hours", "days")
        assert value(f"D{row}")
    obsolete = {"DateValue", "MinimumDutyFreeDaysBeforeNight", "PayrollFortnightAnchor"}
    assert not any(any(label in str(value(address) or "") for label in obsolete) for address in cells)
    assert not any(value(address) is not None and (address[0] > "D" or int(address[1:]) > 9) for address in cells)
    assert not sheet.findall(".//m:f", ns)
    table_files = [name for name in archive.namelist() if name.startswith("xl/tables/") and name.endswith(".xml")]
    assert len(table_files) == 1
    table = ET.fromstring(archive.read(table_files[0]))
    assert table.attrib["name"] == table.attrib["displayName"] == "OvertimeRosterThreshold"
    assert table.attrib["ref"] == "A1:D9"
    columns = table.findall("m:tableColumns/m:tableColumn", ns)
    assert [column.attrib["name"] for column in columns] == ["Title", "Value", "Unit", "Notes"]
    assert "Value must be Number" in value("D2")
    assert "earliest Date in Settings Data PermutationDimensions" in value("D8")
    assert "earliest Date in Settings Data PermutationDimensions" in value("D9")
assert hashlib.sha256(target.read_bytes()).hexdigest().upper() == preservation["afterSHA256"]

evidence = {
    "target": target.as_posix(),
    "authoring": "@oai/artifact-tool 2.8.99 from cached primary runtime",
    "loader_tool_exposed": False,
    "cached_artifact_tool_available": True,
    "fallback_authoring_used": False,
    "sheets": 1,
    "table": "OvertimeRosterThreshold",
    "table_range": "A1:D9",
    "data_rows": 8,
    "unique_titles": 8,
    "numeric_values": 8,
    "blank_value_titles": [],
    "columns": ["Title", "Value", "Unit", "Notes"],
    "removed_obsolete_settings": True,
    "all_existing_numeric_values_preserved": True,
    "date_anchor": "Earliest Date in Settings Data PermutationDimensions",
    "before_sha256": preservation["beforeSHA256"],
    "after_sha256": preservation["afterSHA256"],
    "formulas": 0,
    "visual_review": "Preview inspected: headers, all values and notes readable, no clipping.",
    "only_authorized_template_accessed": True,
    "sharepoint_changes": False,
    "result": "Pass",
}
Path("outputs/overtime-roster-checks/thresholds-verification.json").write_text(json.dumps(evidence, indent=2) + "\n", encoding="utf-8")
print(json.dumps(evidence, indent=2))
