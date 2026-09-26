import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const sourcePath = "C:/taully_demand_forecast/data/historial_demanda.csv";
const outputDir = "C:/taully_demand_forecast/outputs/01a0b651-7620-73f0-8f6a-8b68b14f7a17";
const outputPath = `${outputDir}/inventario_diario_simulado_desde_reportes_2026-04-01_a_2026-09-15.xlsx`;

function parseDate(dateText) {
  const [year, month, day] = dateText.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function parseHistory(csvText) {
  return csvText.trim().split(/\r?\n/).slice(1).map((line) => {
    const [date, category, quantity] = line.split(",");
    return { date, category, quantity: Number(quantity) };
  });
}

const csvText = await fs.readFile(sourcePath, "utf8");
const history = parseHistory(csvText).sort((a, b) => a.date.localeCompare(b.date) || a.category.localeCompare(b.category));
const categories = [...new Set(history.map((record) => record.category))].sort();
const dates = [...new Set(history.map((record) => record.date))].sort();

const recordsByCategory = new Map(categories.map((category) => [category, history.filter((record) => record.category === category)]));
const assumptions = new Map();
const inventoryRows = [];
const summaryRows = [];

for (const category of categories) {
  const records = recordsByCategory.get(category);
  const totalSales = records.reduce((sum, record) => sum + record.quantity, 0);
  const averageDailySales = totalSales / records.length;
  const targetStock = Math.ceil(averageDailySales * 10);
  const reorderPoint = Math.ceil(averageDailySales * 3);
  assumptions.set(category, { averageDailySales, targetStock, reorderPoint });

  let openingStock = targetStock;
  let totalEntries = 0;
  let daysWithoutStock = 0;

  for (const record of records) {
    const entry = openingStock <= reorderPoint ? targetStock - openingStock : 0;
    const availableStock = openingStock + entry;
    const stockoutDay = record.quantity > availableStock ? 1 : 0;
    const closingStock = Math.max(0, availableStock - record.quantity);

    totalEntries += entry;
    daysWithoutStock += stockoutDay;
    inventoryRows.push([
      parseDate(record.date),
      category,
      record.quantity,
      openingStock,
      entry,
      availableStock,
      closingStock,
      stockoutDay,
      reorderPoint,
      targetStock,
      "Historial de ventas real",
      "Simulado con regla de reposición",
    ]);
    openingStock = closingStock;
  }

  summaryRows.push([
    category,
    totalSales,
    targetStock,
    totalEntries,
    openingStock,
    daysWithoutStock,
    records.length,
    null,
    null,
  ]);
}

const actualTotalSales = history.reduce((sum, record) => sum + record.quantity, 0);
const simulatedInitialStock = summaryRows.reduce((sum, row) => sum + row[2], 0);
const simulatedEntries = summaryRows.reduce((sum, row) => sum + row[3], 0);
const simulatedFinalStock = summaryRows.reduce((sum, row) => sum + row[4], 0);
const simulatedStockoutDays = summaryRows.reduce((sum, row) => sum + row[5], 0);
const observedDays = summaryRows.reduce((sum, row) => sum + row[6], 0);

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Resumen");
const daily = workbook.worksheets.add("Inventario diario");

const colors = {
  navy: "#173F5F",
  blue: "#245580",
  paleBlue: "#EAF2F8",
  amber: "#FFF2CC",
  amberDark: "#BF9000",
  paleGray: "#F4F6F8",
  green: "#E2F0D9",
  border: "#D9E2F3",
  text: "#1F2937",
  note: "#5B677A",
};
const bodyFont = { name: "Arial", size: 10, color: colors.text };
const titleFont = { name: "Arial", size: 16, bold: true, color: colors.navy };
const headerFont = { name: "Arial", size: 10, bold: true, color: "#FFFFFF" };

for (const sheet of [summary, daily]) {
  sheet.showGridLines = false;
  sheet.getRange("A1:L900").format.font = bodyFont;
  sheet.getRange("A1:L900").format.verticalAlignment = "center";
}

summary.getRange("A2").values = [["Inventario diario simulado desde reportes"]];
summary.getRange("A2").format.font = titleFont;
summary.mergeCells("A3:I3");
summary.getRange("A3").values = [["Periodo de ventas reales: 01/04/2026 al 15/09/2026. Base: 168 reportes, 5 categorías y 840 registros diarios."]];
summary.getRange("A3:I3").format = { font: { name: "Arial", size: 10, color: colors.note, italic: true }, verticalAlignment: "center" };
summary.mergeCells("A4:I4");
summary.getRange("A4").values = [["IMPORTANTE: las ventas y fechas son reales. SI, EN, SF y DQS son datos simulados para completar el instrumento de inventario; no sustituyen registros físicos de la tienda."]];
summary.getRange("A4:I4").format = { fill: colors.amber, font: { name: "Arial", size: 10, bold: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
summary.getRange("A4:I4").format.rowHeight = 30;

summary.getRange("A6:I6").values = [[
  "Categoría",
  "Ventas CD reales",
  "Stock inicial SI simulado",
  "Entradas EN simuladas",
  "Stock final SF simulado",
  "Días sin stock DQS sim.",
  "Días observados DD reales",
  "Índice salida ISI sim.",
  "Tasa quiebre TQS sim.",
]];
summary.getRange("A6:I6").format = { fill: colors.blue, font: headerFont, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
summary.getRange("A6:I12").format.borders = { preset: "all", style: "thin", color: colors.border };
summary.getRange("A7:I11").values = summaryRows;
summary.getRange("H7").formulas = [["=B7/(C7+D7)"]];
summary.getRange("H7:H11").fillDown();
summary.getRange("I7").formulas = [["=F7/G7"]];
summary.getRange("I7:I11").fillDown();
summary.getRange("A12:G12").values = [["TOTAL", actualTotalSales, simulatedInitialStock, simulatedEntries, simulatedFinalStock, simulatedStockoutDays, observedDays]];
summary.getRange("H12").formulas = [["=B12/(C12+D12)"]];
summary.getRange("I12").formulas = [["=F12/G12"]];
summary.getRange("A12:I12").format = { fill: colors.paleBlue, font: { name: "Arial", size: 10, bold: true, color: colors.navy } };
summary.getRange("B7:G12").format.numberFormat = "#,##0";
summary.getRange("H7:I12").format.numberFormat = "0.0%";
summary.getRange("C7:F12").format.fill = colors.amber;
summary.getRange("H7:I12").format.fill = colors.amber;
summary.getRange("B7:B12").format.fill = colors.paleBlue;

summary.mergeCells("A15:I15");
summary.getRange("A15").values = [["Método reproducible: por categoría, el stock objetivo = redondeo hacia arriba de 10 × la venta diaria promedio real. Si el stock inicial del día es igual o menor al punto de reposición (3 × el promedio diario), se registra una entrada simulada hasta el nivel objetivo."]];
summary.getRange("A15:I15").format = { fill: colors.paleGray, font: { name: "Arial", size: 10, color: colors.note }, wrapText: true, verticalAlignment: "center" };
summary.getRange("A15:I15").format.rowHeight = 38;
summary.mergeCells("A17:I17");
summary.getRange("A17").values = [["En este escenario de reposición diaria, la simulación no presenta días sin stock. El cero en DQS/TQS es resultado del modelo simulado, no una afirmación sobre el inventario real de la tienda."]];
summary.getRange("A17:I17").format = { fill: colors.green, font: { name: "Arial", size: 10, color: colors.navy, italic: true }, wrapText: true, verticalAlignment: "center" };
summary.getRange("A17:I17").format.rowHeight = 30;

const summaryWidths = [17, 17, 21, 20, 20, 21, 22, 18, 19];
summaryWidths.forEach((width, index) => summary.getRangeByIndexes(0, index, 20, 1).format.columnWidth = width);
summary.getRange("A6:I12").format.rowHeight = 28;

daily.getRange("A2").values = [["Inventario diario por categoría"]];
daily.getRange("A2").format.font = titleFont;
daily.mergeCells("A3:L3");
daily.getRange("A3").values = [["Ventas y fechas: historial real consolidado de los 168 reportes. Campos de inventario: simulación reproducible para demostración."]];
daily.getRange("A3:L3").format = { font: { name: "Arial", size: 10, color: colors.note, italic: true }, verticalAlignment: "center" };
daily.mergeCells("A4:L4");
daily.getRange("A4").values = [["Fuente: reportes de ventas incorporados al sistema, del 01/04/2026 al 15/09/2026. Esta hoja no modifica las ventas reales ni registra inventario físico real."]];
daily.getRange("A4:L4").format = { fill: colors.amber, font: { name: "Arial", size: 10, bold: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
daily.getRange("A4:L4").format.rowHeight = 28;

daily.getRange("A6:L6").values = [[
  "Fecha real",
  "Categoría",
  "Ventas CD reales",
  "Stock inicial SI simulado",
  "Entradas EN simuladas",
  "Stock disponible sim.",
  "Stock final SF simulado",
  "DQS sim. (0/1)",
  "Punto reposición sim.",
  "Stock objetivo sim.",
  "Origen de ventas",
  "Origen de inventario",
]];
daily.getRange("A6:L6").format = { fill: colors.blue, font: headerFont, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
daily.getRangeByIndexes(6, 0, inventoryRows.length, 12).values = inventoryRows;
daily.getRange(`A7:A${inventoryRows.length + 6}`).format.numberFormat = "dd/mm/yyyy";
daily.getRange(`C7:J${inventoryRows.length + 6}`).format.numberFormat = "#,##0";
daily.getRange(`D7:J${inventoryRows.length + 6}`).format.fill = colors.amber;
daily.getRange(`C7:C${inventoryRows.length + 6}`).format.fill = colors.paleBlue;
daily.getRange(`K7:K${inventoryRows.length + 6}`).format.fill = colors.paleBlue;
daily.getRange(`L7:L${inventoryRows.length + 6}`).format.fill = colors.amber;
daily.getRange(`A6:L${inventoryRows.length + 6}`).format.borders = { preset: "outside", style: "thin", color: colors.border };
daily.getRange(`A7:L${inventoryRows.length + 6}`).format.rowHeight = 18;
daily.tables.add(`A6:L${inventoryRows.length + 6}`, true, "InventarioDiarioTable");
daily.freezePanes.freezeRows(6);
daily.freezePanes.freezeColumns(2);

const dailyWidths = [13, 16, 16, 20, 19, 19, 20, 14, 19, 18, 24, 28];
dailyWidths.forEach((width, index) => daily.getRangeByIndexes(0, index, inventoryRows.length + 7, 1).format.columnWidth = width);

workbook.recalculate();

const summaryCheck = await workbook.inspect({ kind: "table", range: "Resumen!A2:I17", include: "values,formulas", tableMaxRows: 18, tableMaxCols: 9 });
console.log(summaryCheck.ndjson);
const dailyCheck = await workbook.inspect({ kind: "table", range: "Inventario diario!A1:L12", include: "values", tableMaxRows: 12, tableMaxCols: 12 });
console.log(dailyCheck.ndjson);
const errorCheck = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" });
console.log(errorCheck.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const summaryPreview = await workbook.render({ sheetName: "Resumen", range: "A1:I18", scale: 1.35, format: "png" });
await fs.writeFile(`${outputDir}/inventario_diario_resumen_preview.png`, new Uint8Array(await summaryPreview.arrayBuffer()));
const dailyTopPreview = await workbook.render({ sheetName: "Inventario diario", range: "A1:L20", scale: 1.1, format: "png" });
await fs.writeFile(`${outputDir}/inventario_diario_top_preview.png`, new Uint8Array(await dailyTopPreview.arrayBuffer()));
const dailyTailPreview = await workbook.render({ sheetName: "Inventario diario", range: `A828:L${inventoryRows.length + 6}`, scale: 1.1, format: "png" });
await fs.writeFile(`${outputDir}/inventario_diario_tail_preview.png`, new Uint8Array(await dailyTailPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
