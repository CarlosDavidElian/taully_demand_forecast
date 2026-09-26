import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/taully_demand_forecast/outputs/01a0b651-7620-73f0-8f6a-8b68b14f7a17";
const outputPath = `${outputDir}/datos_inventario_simulados_2026-04-01_a_2026-09-15.xlsx`;

const rows = [
  [new Date(2026, 3, 1), new Date(2026, 8, 15), "ABARROTES", 6800, 4700, 8577, null, 2, 168, "Historial de ventas real", "Simulado para demostración"],
  [new Date(2026, 3, 1), new Date(2026, 8, 15), "BEBIDAS", 11400, 8200, 15046, null, 1, 168, "Historial de ventas real", "Simulado para demostración"],
  [new Date(2026, 3, 1), new Date(2026, 8, 15), "GOLOSINAS", 9500, 7000, 12498, null, 3, 168, "Historial de ventas real", "Simulado para demostración"],
  [new Date(2026, 3, 1), new Date(2026, 8, 15), "HELADOS", 3100, 2600, 4200, null, 5, 168, "Historial de ventas real", "Simulado para demostración"],
  [new Date(2026, 3, 1), new Date(2026, 8, 15), "LIMPIEZA", 2200, 1700, 2784, null, 2, 168, "Historial de ventas real", "Simulado para demostración"],
];

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Datos inventario");
sheet.showGridLines = false;

const navy = "#173F5F";
const blue = "#245580";
const amber = "#FFF2CC";
const paleBlue = "#EAF2F8";
const red = "#9C0006";
const paleRed = "#FCE4D6";
const border = "#D9E2F3";
const bodyFont = { name: "Arial", size: 10, color: "#1F2937" };

sheet.getRange("A1:K20").format.font = bodyFont;
sheet.getRange("A1:K20").format.verticalAlignment = "center";
sheet.getRange("A2").values = [["Datos de inventario simulados"]];
sheet.getRange("A2").format.font = { name: "Arial", size: 16, bold: true, color: navy };
sheet.mergeCells("A3:K3");
sheet.getRange("A3").values = [["SIMULACIÓN PARA PRUEBA. SI, EN, SF y DQS no son datos reales de la tienda. No usar como resultado real de la tesis."]];
sheet.getRange("A3:K3").format = { fill: amber, font: { name: "Arial", size: 10, bold: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
sheet.getRange("A3:K3").format.rowHeight = 28;
sheet.mergeCells("A4:K4");
sheet.getRange("A4").values = [["Ventas CD y días DD: historial consolidado real del 01/04/2026 al 15/09/2026. Stock final SF se calcula como SI + EN - CD."]];
sheet.getRange("A4:K4").format = { fill: paleBlue, font: { name: "Arial", size: 10, italic: true, color: navy }, wrapText: true, verticalAlignment: "center" };
sheet.getRange("A4:K4").format.rowHeight = 24;

sheet.getRange("A6:K6").values = [[
  "Periodo inicio",
  "Periodo fin",
  "Categoría",
  "Stock inicial SI simulado",
  "Entradas EN simuladas",
  "Ventas CD reales",
  "Stock final SF calculado",
  "Días sin stock DQS simulados",
  "Días observados DD reales",
  "Origen de ventas y días",
  "Origen de inventario",
]];
sheet.getRange("A6:K6").format = { fill: blue, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
sheet.getRange("A6:K11").format.borders = { preset: "all", style: "thin", color: border };
sheet.getRange("A7:K11").values = rows;
sheet.getRange("G7:G11").formulas = rows.map((_, index) => [`=D${index + 7}+E${index + 7}-F${index + 7}`]);
sheet.getRange("A7:B11").format.numberFormat = "dd/mm/yyyy";
sheet.getRange("D7:I11").format.numberFormat = "#,##0";
sheet.getRange("D7:E11").format.fill = amber;
sheet.getRange("H7:H11").format.fill = amber;
sheet.getRange("D6:E6").format.fill = "#BF9000";
sheet.getRange("H6:H6").format.fill = "#BF9000";

sheet.mergeCells("A14:K15");
sheet.getRange("A14").values = [["Uso de los campos: ISI = CD / (SI + EN). TQS = DQS / DD. Este archivo contiene los campos que faltaban para realizar ambos cálculos, pero sus valores de inventario son simulados."]];
sheet.getRange("A14:K15").format = { fill: paleRed, font: { name: "Arial", size: 10, color: red }, wrapText: true, verticalAlignment: "center" };
sheet.getRange("A14:K15").format.rowHeight = 22;

const widths = [14, 14, 17, 21, 20, 16, 21, 24, 22, 24, 26];
widths.forEach((width, index) => sheet.getRangeByIndexes(0, index, 20, 1).format.columnWidth = width);
sheet.getRange("A6:K11").format.rowHeight = 26;
sheet.freezePanes.freezeRows(6);

workbook.recalculate();
const verification = await workbook.inspect({ kind: "table", range: "Datos inventario!A2:K15", include: "values,formulas", tableMaxRows: 16, tableMaxCols: 11 });
console.log(verification.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "Datos inventario", range: "A1:K15", scale: 1.35, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(`${outputDir}/datos_inventario_simulados_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
