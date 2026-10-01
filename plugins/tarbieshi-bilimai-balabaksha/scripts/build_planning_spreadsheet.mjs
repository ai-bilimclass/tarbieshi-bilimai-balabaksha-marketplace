#!/usr/bin/env node
// Export the content of a completed planning DOCX as an editable Excel table.
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [inputPath, outputPath, pythonPath, landscapeScript] = process.argv.slice(2);
if (!inputPath || !outputPath) {
  throw new Error("Usage: node build_planning_spreadsheet.mjs payload.json output.xlsx");
}
const data = JSON.parse(await fs.readFile(inputPath, "utf8"));
const expectedColumns = { annual: 3, weekly: 6, daily: 2 }[data.mode];
if (!expectedColumns || !Array.isArray(data.rows) || data.rows.length < 2 ||
    data.rows.some(row => !Array.isArray(row) || row.length !== expectedColumns)) {
  throw new Error("Invalid planning table payload");
}
if ((data.mode === "annual" || data.mode === "weekly") && (!pythonPath || !landscapeScript)) {
  throw new Error("Annual and weekly XLSX exports require Python and set_xlsx_landscape.py paths");
}

const workbook = Workbook.create();
const rawName = String(data.prefix?.[0] || "Plan").replace(/[\\/*?:\[\]]/g, " ").trim();
const sheet = workbook.worksheets.add(rawName.slice(0, 31) || "Plan");
sheet.showGridLines = false;
const columns = expectedColumns;
const prefix = data.prefix || [];
const suffix = data.suffix || [];
const tableStart = prefix.length + 2; // one spacer row before the table
const tableEnd = tableStart + data.rows.length - 1;
const totalRows = tableEnd + suffix.length + 1;
const all = sheet.getRangeByIndexes(0, 0, totalRows, columns);
all.format.font = { name: "Times New Roman", size: 12, color: "#000000" };
all.format.wrapText = true;
all.format.verticalAlignment = "top";

const widths = data.mode === "weekly" ? [38, 64, 64, 64, 64, 64] :
               data.mode === "annual" ? [21, 40, 110] : [44, 110];
widths.forEach((width, index) => {
  sheet.getRangeByIndexes(0, index, totalRows, 1).format.columnWidth = width;
});
const safeText = value => {
  const text = String(value ?? "");
  return text.startsWith("=") ? `'${text}` : text;
};
const writeWide = (rowIndex, value, bold = false) => {
  const range = sheet.getRangeByIndexes(rowIndex, 0, 1, columns);
  range.merge();
  range.values = [[safeText(value)]];
  range.format.font = { name: "Times New Roman", size: 12, color: "#000000", bold };
  range.format.rowHeight = Math.min(150, Math.max(26, Math.ceil(String(value).length / 110) * 18));
};
prefix.forEach((value, index) => writeWide(index, value, index === 0));

const table = sheet.getRangeByIndexes(tableStart - 1, 0, data.rows.length, columns);
table.values = data.rows.map(row => row.map(safeText));
table.format.borders = { preset: "all", style: "thin", color: "#000000" };
const header = sheet.getRangeByIndexes(tableStart - 1, 0, 1, columns);
header.format.font = { name: "Times New Roman", size: 12, color: "#000000", bold: true };
header.format.rowHeight = 42;
for (let index = 1; index < data.rows.length; index++) {
  const row = data.rows[index];
  const maxLines = Math.max(...row.map((text, col) =>
    String(text).split("\n").reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / widths[col])), 0)));
  sheet.getRangeByIndexes(tableStart - 1 + index, 0, 1, columns).format.rowHeight =
    Math.min(390, Math.max(38, maxLines * 17));
}
if (data.mode === "annual") {
  for (const [first, last] of data.month_merges || []) {
    if (last > first) {
      sheet.getRangeByIndexes(tableStart - 1 + first, 0, last - first + 1, 1).merge();
    }
  }
}
suffix.forEach((value, index) => writeWide(tableEnd + 1 + index, value));
sheet.freezePanes.freezeRows(tableStart);

workbook.recalculate();
const check = await workbook.inspect({ kind: "region", sheetId: sheet.name,
  range: `A${tableStart}:${String.fromCharCode(64 + columns)}${Math.min(tableStart + 2, tableEnd)}`,
  maxChars: 1200 });
if (!check?.ndjson) throw new Error("Workbook inspection failed");
await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
if (data.mode === "annual" || data.mode === "weekly") {
  const result = spawnSync(pythonPath, [landscapeScript, outputPath], { encoding: "utf8" });
  if (result.error || result.status !== 0) {
    throw new Error(`Landscape XLSX setup failed: ${result.error?.message || result.stderr || result.stdout}`);
  }
}
console.log(`Saved ${outputPath}: ${data.rows.length - 1} data rows, ${columns} columns`);
