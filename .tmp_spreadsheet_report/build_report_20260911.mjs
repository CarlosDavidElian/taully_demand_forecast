import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const taskDir = process.cwd();
const sourcePath = path.resolve(taskDir, "..", "data", "reporte_20260429.xlsx");
const outputDir = path.resolve(taskDir, "..", "outputs", "reporte_20260911");
const outputPath = path.join(outputDir, "reporte_20260911.xlsx");

const source = await FileBlob.load(sourcePath);
const workbook = await SpreadsheetFile.importXlsx(source);
const sheet = workbook.worksheets.getItem("Hoja1");

sheet.getRange("A3:B3").values = [["FECHAI: 11/09/2026", "FECHAF: 11/09/2026"]];
workbook.recalculate();

const validation = await workbook.inspect({
  kind: "table",
  range: "Hoja1!A1:D15",
  include: "values,formulas",
  tableMaxRows: 15,
  tableMaxCols: 4,
});
console.log(validation.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({
  sheetName: "Hoja1",
  range: "A1:D15",
  scale: 1.5,
  format: "png",
});
await fs.writeFile(path.join(outputDir, "preview.png"), new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Created ${outputPath}`);
