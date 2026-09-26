import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/taully_demand_forecast/outputs/01a0b651-7620-73f0-8f6a-8b68b14f7a17";
const outputPath = `${outputDir}/demostracion_indicadores_inventario_simulados.xlsx`;

const data = [
  ["ABARROTES", 8577, 168, 6800, 4700, 2, "Ventas y días: historial real. Inventario: simulación."],
  ["BEBIDAS", 15046, 168, 11400, 8200, 1, "Ventas y días: historial real. Inventario: simulación."],
  ["GOLOSINAS", 12498, 168, 9500, 7000, 3, "Ventas y días: historial real. Inventario: simulación."],
  ["HELADOS", 4200, 168, 3100, 2600, 5, "Ventas y días: historial real. Inventario: simulación."],
  ["LIMPIEZA", 2784, 168, 2200, 1700, 2, "Ventas y días: historial real. Inventario: simulación."],
];

const workbook = Workbook.create();
const result = workbook.worksheets.add("Resultados demo");
const inputs = workbook.worksheets.add("Datos simulados");

const navy = "#173F5F";
const blue = "#245580";
const paleBlue = "#EAF2F8";
const paleAmber = "#FFF2CC";
const paleRed = "#FCE4D6";
const darkText = "#1F2937";
const border = "#D9E2F3";
const font = { name: "Arial", size: 10, color: darkText };
const header = { fill: blue, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };

for (const sheet of [result, inputs]) {
  sheet.showGridLines = false;
  sheet.getRange("A1:H30").format.font = font;
  sheet.getRange("A1:H30").format.verticalAlignment = "center";
}

// Hoja de resultados
result.getRange("A2").values = [["Demostración de indicadores de inventario"]];
result.getRange("A2").format.font = { name: "Arial", size: 16, bold: true, color: navy };
result.mergeCells("A3:H3");
result.getRange("A3").values = [["SIMULACIÓN PARA DEMOSTRACIÓN. Los valores de stock, entradas y quiebres no son registros reales de la tienda."]];
result.getRange("A3:H3").format = { fill: paleAmber, font: { name: "Arial", size: 10, bold: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
result.getRange("A3:H3").format.rowHeight = 30;

result.getRange("A5:B8").values = [
  ["Ventas reales de referencia", null],
  ["Índice de salida simulado", null],
  ["Tasa de quiebre simulada", null],
  ["Cobertura real", "5 categorías y 168 días"],
];
result.getRange("A5:A8").format = { fill: paleBlue, font: { name: "Arial", size: 10, bold: true, color: navy }, verticalAlignment: "center" };
result.getRange("A5:B8").format.borders = { preset: "all", style: "thin", color: border };
result.getRange("B5").formulas = [["=SUM(B11:B15)"]];
result.getRange("B6").formulas = [["=SUM(B11:B15)/(SUM(C11:C15)+SUM(D11:D15))"]];
result.getRange("B7").formulas = [["=SUM(F11:F15)/SUM(G11:G15)"]];
result.getRange("B5").format.numberFormat = "#,##0";
result.getRange("B6:B7").format.numberFormat = "0.00%";

result.getRange("A10:H10").values = [["Categoría", "CD ventas reales", "SI simulado", "EN simulado", "ISI simulado", "DQS simulado", "DD real", "TQS simulada"]];
result.getRange("A10:H10").format = header;
result.getRange("A10:H15").format.borders = { preset: "all", style: "thin", color: border };
result.getRange("A11:A15").formulas = data.map((_, index) => [`='Datos simulados'!A${index + 6}`]);
result.getRange("B11:B15").formulas = data.map((_, index) => [`='Datos simulados'!B${index + 6}`]);
result.getRange("C11:C15").formulas = data.map((_, index) => [`='Datos simulados'!D${index + 6}`]);
result.getRange("D11:D15").formulas = data.map((_, index) => [`='Datos simulados'!E${index + 6}`]);
result.getRange("E11:E15").formulas = data.map((_, index) => [`=B${index + 11}/(C${index + 11}+D${index + 11})`]);
result.getRange("F11:F15").formulas = data.map((_, index) => [`='Datos simulados'!F${index + 6}`]);
result.getRange("G11:G15").formulas = data.map((_, index) => [`='Datos simulados'!C${index + 6}`]);
result.getRange("H11:H15").formulas = data.map((_, index) => [`=F${index + 11}/G${index + 11}`]);
result.getRange("B11:D15").format.numberFormat = "#,##0";
result.getRange("E11:E15").format.numberFormat = "0.00%";
result.getRange("F11:G15").format.numberFormat = "#,##0";
result.getRange("H11:H15").format.numberFormat = "0.00%";
result.getRange("B11:H15").format.horizontalAlignment = "right";

result.getRange("A18").values = [["Cómo se calcula"]];
result.getRange("A18").format.font = { name: "Arial", size: 11, bold: true, color: navy };
result.mergeCells("A19:H20");
result.getRange("A19").values = [["ISI = CD / (SI + EN). TQS = DQS / DD. CD y DD provienen del historial de ventas real; SI, EN y DQS son solo valores simulados para mostrar el cálculo."]];
result.getRange("A19:H20").format = { fill: paleRed, font: { name: "Arial", size: 10, color: "#9C0006" }, wrapText: true, verticalAlignment: "center" };
result.getRange("A19:H20").format.rowHeight = 22;

// Hoja de datos
inputs.getRange("A2").values = [["Datos de referencia y valores simulados"]];
inputs.getRange("A2").format.font = { name: "Arial", size: 14, bold: true, color: navy };
inputs.mergeCells("A3:G3");
inputs.getRange("A3").values = [["Las ventas y los días son datos reales del historial. Las columnas de inventario son simuladas y únicamente sirven para probar las fórmulas."]];
inputs.getRange("A3:G3").format = { fill: paleAmber, font: { name: "Arial", size: 10, bold: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
inputs.getRange("A3:G3").format.rowHeight = 30;
inputs.getRange("A5:G5").values = [["Categoría", "Ventas históricas reales", "Días registrados reales", "Stock inicial simulado", "Entradas simuladas", "Días sin stock simulados", "Origen y alcance"]];
inputs.getRange("A5:G5").format = header;
inputs.getRange("A6:G10").values = data;
inputs.getRange("A5:G10").format.borders = { preset: "all", style: "thin", color: border };
inputs.getRange("B6:F10").format.numberFormat = "#,##0";
inputs.getRange("D6:F10").format.fill = paleAmber;
inputs.getRange("D5:F5").format.fill = "#BF9000";
inputs.mergeCells("A13:G14");
inputs.getRange("A13").values = [["Método de demostración: las ventas reales se tomaron del historial consolidado del 01/04/2026 al 15/09/2026. Los valores de inventario fueron elegidos solo para demostrar el funcionamiento de ISI y TQS."]];
inputs.getRange("A13:G14").format = { fill: paleRed, font: { name: "Arial", size: 10, color: "#9C0006" }, wrapText: true, verticalAlignment: "center" };
inputs.getRange("A13:G14").format.rowHeight = 22;

for (const [sheet, widths] of [
  [result, [28, 20, 16, 16, 15, 16, 14, 16]],
  [inputs, [20, 22, 20, 20, 18, 22, 42]],
]) {
  widths.forEach((width, index) => sheet.getRangeByIndexes(0, index, 30, 1).format.columnWidth = width);
}
result.getRange("A10:H15").format.rowHeight = 23;
inputs.getRange("A5:G10").format.rowHeight = 23;

result.freezePanes.freezeRows(10);
inputs.freezePanes.freezeRows(5);

workbook.recalculate();

const resultCheck = await workbook.inspect({ kind: "table", range: "Resultados demo!A2:H19", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 8 });
console.log(resultCheck.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" });
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: "Resultados demo", range: "A1:H20", scale: 1.5, format: "png" });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(`${outputDir}/demostracion_indicadores_inventario_simulados_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const inputsPreview = await workbook.render({ sheetName: "Datos simulados", range: "A1:G14", scale: 1.5, format: "png" });
await fs.writeFile(`${outputDir}/demostracion_indicadores_inventario_simulados_datos_preview.png`, new Uint8Array(await inputsPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
