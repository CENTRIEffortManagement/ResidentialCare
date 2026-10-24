"""Read only the authorized BD RN B saved values and metadata."""
import hashlib
import importlib.util
import json
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from datetime import datetime, timezone

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("saved_tables", ROOT / "outputs/jh-ry-ain-availability-20261021/inspect_saved_tables.py")
helper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(helper)
RELATIVE = "CLIENT/DATExx-Whiddon/UNITS/BD/2. Calculations/RN/CapacityDistrib(B)-shifts.xlsx"
path = ROOT / RELATIVE
before = hashlib.sha256(path.read_bytes()).hexdigest()
book = helper.read_book(RELATIVE, {"C__TABLE", "C___TABLE_B", "MaxAvailability", "CapacityDistribB_AVAILABILITY_LINEAGE_CHECK"})
result = {"path": RELATIVE, "sha256_before": before,
          "last_write_utc": datetime.fromtimestamp(path.stat().st_mtime, timezone.utc).isoformat(),
          "tables": []}
for table in book["tables"]:
    if "data" not in table:
        continue
    selected = [{"excel_row": index + 2, **row} for index, row in enumerate(table["data"]) if row.get("Resource") == 75]
    metadata = {key: value for key, value in table.items() if key != "data"}
    metadata["resource75"] = selected
    if selected:
        amount = "C##" if table["name"] == "C__TABLE" else "C###"
        metadata["sum"] = sum(row[amount] or 0 for row in selected)
        metadata["positive_periods"] = [row["Period"] for row in selected if row[amount] is not None and row[amount] > 0]
        metadata["audit_values"] = {col: sorted(set(row[col] for row in selected)) for col in ["ResAvailability", "ResMaxAvail"] if col in selected[0]}
    else:
        metadata["data"] = table["data"]
    result["tables"].append(metadata)

with ZipFile(path) as archive:
    strings = []
    if "xl/sharedStrings.xml" in archive.namelist():
        strings = ["".join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
    def decode(cell):
        kind = cell.get("t")
        value = cell.find("s:v", helper.NS)
        if kind == "inlineStr":
            value = "".join(cell.find("s:is", helper.NS).itertext())
        elif value is None:
            value = None
        elif kind == "s":
            value = strings[int(value.text)]
        else:
            value = value.text
        formula = cell.find("s:f", helper.NS)
        return {"address": cell.attrib["r"], "value": value, "kind": kind,
                "formula": None if formula is None else formula.text}
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = helper.xml_relationships(archive, "xl/workbook.xml")
    result["sheets"] = []
    for sheet in workbook.find("s:sheets", helper.NS):
        part = relationships[sheet.attrib[f"{{{helper.REL}}}id"]]
        xml = ET.fromstring(archive.read(part))
        item = {"name": sheet.attrib["name"], "dimension": xml.find("s:dimension", helper.NS).attrib["ref"]}
        if sheet.attrib["name"] == "Sheet1":
            item["cells"] = [decode(cell) for row in xml.findall("s:sheetData/s:row", helper.NS) for cell in row]
        result["sheets"].append(item)
    result["core_properties_xml"] = archive.read("docProps/core.xml").decode()
    result["query_table_metadata"] = [
        {"part": part, "xml": archive.read(part).decode()} for part in archive.namelist()
        if part.startswith("xl/queryTables/") and part.endswith(".xml")]
result["sha256_after"] = hashlib.sha256(path.read_bytes()).hexdigest()
result["unchanged"] = result["sha256_after"] == before
(ROOT / "outputs/bd-rn-resource75-20261021/B-RN-resource75-evidence.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
print(json.dumps({**{key: value for key, value in result.items() if key not in {"tables", "core_properties_xml", "query_table_metadata"}},
                  "tables": [{key: value for key, value in table.items() if key != "resource75"} for table in result["tables"]]}, indent=2))
