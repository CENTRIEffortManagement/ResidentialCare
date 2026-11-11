import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

// The workbook is an explicitly requested SharePoint import template.
// Amend only this requested template, preserving its eight retained numeric inputs.
// Paths remain relative to the repository. No other workbook is read.
const evidenceDir = path.resolve("outputs/overtime-roster-checks");
const outputPath = path.resolve("docs/OvertimeRosterCheck/OvertimeRosterThreshold.xlsx");
const retainedTitles = [
  "DailyHoursLimit", "CSHospitalityHoursLimit", "CSPersonalCareHoursLimit", "MinimumBreakHours",
  "ShortGapReviewHours", "MinimumFortnightDutyFreeDays", "FortnightHoursLimit", "FortnightLengthDays",
];

await fs.mkdir(evidenceDir, { recursive: true });
await fs.mkdir(path.dirname(outputPath), { recursive: true });
const beforeBytes = await fs.readFile(outputPath);
const beforeSHA256 = crypto.createHash("sha256").update(beforeBytes).digest("hex").toUpperCase();
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const sheet = workbook.worksheets.getItem("OvertimeRosterThreshold");
const previousRows = sheet.getRange("A1:E11").values;
const headings = previousRows[0];
const indexes = Object.fromEntries(["Title", "Value", "Unit", "Notes"].map((name) => [name, headings.indexOf(name)]));
if (Object.values(indexes).some((index) => index < 0)) throw new Error("Template input columns are incomplete.");
const retainedInputs = retainedTitles.map((title) => {
  const matches = previousRows.slice(1).filter((row) => row[indexes.Title] === title);
  if (matches.length !== 1) throw new Error(`Expected exactly one retained setting: ${title}`);
  const input = matches[0];
  if (typeof input[indexes.Value] !== "number" || !Number.isFinite(input[indexes.Value])) {
    throw new Error(`Retained setting must remain numeric: ${title}`);
  }
  return [title, input[indexes.Value], input[indexes.Unit], input[indexes.Notes]];
});
const rows = [["Title", "Value", "Unit", "Notes"], ...retainedInputs.map((row) => {
  if (row[0] === "FortnightHoursLimit") return [...row.slice(0, 3), "Count all net rostered hours once, including daily excess. Anchor fortnights to the earliest Date in Settings Data PermutationDimensions."];
  if (row[0] === "FortnightLengthDays") return [...row.slice(0, 3), "Fortnight length. Date anchor is the earliest Date in Settings Data PermutationDimensions."];
  return row;
})];
const sourceTables = sheet.tables.items;
if (sourceTables.length !== 1 || sourceTables[0].name !== "OvertimeRosterThreshold") throw new Error("Unexpected template table contract.");
const tableStyle = sourceTables[0].style;
sourceTables[0].delete();
sheet.getRange("A1:E11").clear({ applyTo: "all" });
sheet.getRange("A1:D9").values = rows;
sheet.getRange("A1:D9").format.font = { name: "Arial", size: 10, color: "#202833" };
sheet.getRange("A1:D9").format.verticalAlignment = "center";
sheet.getRange("A1:D9").format.rowHeight = 34;
sheet.getRange("A1:D1").format = {
  fill: "#344054",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  rowHeight: 26,
};
sheet.getRange("A1:A9").format.columnWidth = 40;
sheet.getRange("B1:B9").format.columnWidth = 10;
sheet.getRange("C1:C9").format.columnWidth = 10;
sheet.getRange("D1:D9").format.columnWidth = 90;
sheet.getRange("E1:E11").format.columnWidth = 8;
sheet.getRange("A2:A9").format.horizontalAlignment = "left";
sheet.getRange("B2:B9").format.horizontalAlignment = "right";
sheet.getRange("B2:B9").setNumberFormat("0");
sheet.getRange("C2:D9").format.horizontalAlignment = "left";
sheet.getRange("D2:D9").format.wrapText = true;
const table = sheet.tables.add("A1:D9", true, "OvertimeRosterThreshold");
table.style = tableStyle;
table.showTotals = false;
table.showFilterButton = true;

workbook.recalculate();
const inspect = await workbook.inspect({
  kind: "table",
  range: "OvertimeRosterThreshold!A1:D9",
  include: "values,formulas",
  tableMaxRows: 9,
  tableMaxCols: 4,
  maxChars: 8000,
});
await fs.writeFile(path.join(evidenceDir, "thresholds-inspect.ndjson"), inspect.ndjson);
const preview = await workbook.render({
  sheetName: "OvertimeRosterThreshold",
  range: "A1:D9",
  scale: 1.25,
  format: "png",
});
await fs.writeFile(path.join(evidenceDir, "thresholds-preview.png"), new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
const afterSHA256 = crypto.createHash("sha256").update(await fs.readFile(outputPath)).digest("hex").toUpperCase();
await fs.writeFile(path.join(evidenceDir, "thresholds-amendment-preservation.json"), JSON.stringify({
  beforeSHA256, afterSHA256,
  retainedValues: Object.fromEntries(retainedInputs.map((row) => [row[0], row[1]])),
  retainedValueCount: retainedInputs.length,
  removedColumns: ["DateValue"],
  removedTitles: ["MinimumDutyFreeDaysBeforeNight", "PayrollFortnightAnchor"],
  anchor: "Earliest Date in Settings Data PermutationDimensions",
  existingNumericValuesPreserved: true,
}, null, 2) + "\n");
// Keep Artifact Tool's export inspection beside the other support evidence.
try {
  await fs.rename(`${outputPath}.inspect.ndjson`, path.join(evidenceDir, "thresholds-export-inspect.ndjson"));
} catch (error) {
  if (error.code !== "ENOENT") throw error;
}
console.log(JSON.stringify({ output: path.relative(process.cwd(), outputPath), rows: rows.length - 1, table: table.name }));
