import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = "C:/taully_demand_forecast/.tmp_inventory_demo/product_sales_from_reports.json";
const outputDir = "C:/taully_demand_forecast/outputs/01a0b651-7620-73f0-8f6a-8b68b14f7a17";
const outputPath = `${outputDir}/inventario_por_producto_y_marca_2026-04-01_a_2026-09-15.xlsx`;

function parseDate(dateText) {
  const [year, month, day] = dateText.split("-").map(Number);
  return new Date(year, month - 1, day);
}

const input = JSON.parse(await fs.readFile(inputPath, "utf8"));
const sales = input.records;
const metadata = input.metadata;
const dates = Array.from({ length: metadata.date_count }, (_, index) => {
  const start = parseDate(metadata.start_date);
  return new Date(start.getFullYear(), start.getMonth(), start.getDate() + index);
});
const dateKeys = dates.map((date) => date.toISOString().slice(0, 10));

const productMap = new Map();
const salesByDateProduct = new Map();
for (const sale of sales) {
  productMap.set(sale.product, {
    product: sale.product,
    brand: sale.brand,
    family: sale.family,
    category: sale.category,
  });
  salesByDateProduct.set(`${sale.date}|${sale.product}`, sale.quantity);
}
const products = [...productMap.values()].sort((a, b) =>
  a.category.localeCompare(b.category) || a.brand.localeCompare(b.brand) || a.product.localeCompare(b.product),
);

const inventoryRows = [];
const productSummaryRows = [];
for (const product of products) {
  const productSales = dateKeys.map((date) => salesByDateProduct.get(`${date}|${product.product}`) ?? 0);
  const totalSales = productSales.reduce((sum, quantity) => sum + quantity, 0);
  const daysWithSales = productSales.filter((quantity) => quantity > 0).length;
  const averageDailySales = totalSales / dateKeys.length;
  const highestDailySale = Math.max(...productSales);
  const targetStock = Math.max(Math.ceil(averageDailySales * 10), highestDailySale * 2, 1);
  const reorderPoint = Math.max(Math.ceil(averageDailySales * 3), 1);
  let openingStock = targetStock;
  let totalEntries = 0;
  let daysWithoutStock = 0;

  for (let index = 0; index < dateKeys.length; index += 1) {
    const quantity = productSales[index];
    const entry = openingStock <= reorderPoint ? targetStock - openingStock : 0;
    const availableStock = openingStock + entry;
    const stockoutDay = quantity > availableStock ? 1 : 0;
    const closingStock = Math.max(0, availableStock - quantity);

    totalEntries += entry;
    daysWithoutStock += stockoutDay;
    if (quantity > 0) {
      inventoryRows.push([
        dates[index],
        product.product,
        product.brand,
        product.family,
        product.category,
        quantity,
        openingStock,
        entry,
        availableStock,
        closingStock,
        stockoutDay,
        reorderPoint,
        targetStock,
        "Reporte + catálogo",
        "Simulado",
      ]);
    }
    openingStock = closingStock;
  }

  productSummaryRows.push([
    product.product,
    product.brand,
    product.family,
    product.category,
    daysWithSales,
    totalSales,
    targetStock,
    totalEntries,
    openingStock,
    daysWithoutStock,
    dateKeys.length,
    null,
    null,
  ]);
}

const totalActualSales = productSummaryRows.reduce((sum, row) => sum + row[5], 0);
const totalInitialStock = productSummaryRows.reduce((sum, row) => sum + row[6], 0);
const totalEntries = productSummaryRows.reduce((sum, row) => sum + row[7], 0);
const totalFinalStock = productSummaryRows.reduce((sum, row) => sum + row[8], 0);
const totalStockoutDays = productSummaryRows.reduce((sum, row) => sum + row[9], 0);
const totalObservedDays = productSummaryRows.reduce((sum, row) => sum + row[10], 0);
const brands = new Set(products.map((product) => product.brand));

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Resumen por producto");
const daily = workbook.worksheets.add("Inventario diario");

const colors = {
  navy: "#173F5F",
  blue: "#245580",
  paleBlue: "#EAF2F8",
  amber: "#FFF2CC",
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
}

summary.getRange("A1:M210").format.font = bodyFont;
summary.getRange("A1:M210").format.verticalAlignment = "center";
summary.getRange("A2").values = [["Inventario por producto y marca"]];
summary.getRange("A2").format.font = titleFont;
summary.mergeCells("A3:M3");
summary.getRange("A3").values = [[`Ventas reales de ${metadata.report_count} reportes, del 01/04/2026 al 15/09/2026. Se identificaron ${products.length} productos vendidos y ${brands.size} marcas.`]];
summary.getRange("A3:M3").format = { font: { name: "Arial", size: 10, color: colors.note, italic: true }, verticalAlignment: "center" };
summary.mergeCells("A4:M4");
summary.getRange("A4").values = [["IMPORTANTE: producto, marca, categoría, familia y ventas son datos reales de los reportes y del catálogo. SI, EN, SF y DQS son una simulación de inventario para el instrumento de tesis."]];
summary.getRange("A4:M4").format = { fill: colors.amber, font: { name: "Arial", size: 10, bold: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
summary.getRange("A4:M4").format.rowHeight = 30;

summary.getRange("A6:M6").values = [[
  "Producto vendido",
  "Marca",
  "Familia",
  "Categoría",
  "Días con venta real",
  "Ventas CD reales",
  "Stock inicial SI sim.",
  "Entradas EN sim.",
  "Stock final SF sim.",
  "DQS sim.",
  "Días historial",
  "ISI sim.",
  "TQS sim.",
]];
summary.getRange("A6:M6").format = { fill: colors.blue, font: headerFont, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
summary.getRange(`A7:M${products.length + 6}`).values = productSummaryRows;
summary.getRange("L7").formulas = [["=F7/(G7+H7)"]];
summary.getRange(`L7:L${products.length + 6}`).fillDown();
summary.getRange("M7").formulas = [["=J7/K7"]];
summary.getRange(`M7:M${products.length + 6}`).fillDown();
const totalRow = products.length + 7;
summary.getRange(`A${totalRow}:K${totalRow}`).values = [[`TOTAL (${products.length} productos)`, `${brands.size} marcas`, null, null, null, totalActualSales, totalInitialStock, totalEntries, totalFinalStock, totalStockoutDays, totalObservedDays]];
summary.getRange(`L${totalRow}`).formulas = [[`=F${totalRow}/(G${totalRow}+H${totalRow})`]];
summary.getRange(`M${totalRow}`).formulas = [[`=J${totalRow}/K${totalRow}`]];
summary.getRange(`A6:M${totalRow}`).format.borders = { preset: "outside", style: "thin", color: colors.border };
summary.getRange(`F7:K${totalRow}`).format.numberFormat = "#,##0";
summary.getRange(`L7:M${totalRow}`).format.numberFormat = "0.0%";
summary.getRange(`F7:F${totalRow}`).format.fill = colors.paleBlue;
summary.getRange(`G7:J${totalRow}`).format.fill = colors.amber;
summary.getRange(`L7:M${totalRow}`).format.fill = colors.amber;
summary.getRange(`A${totalRow}:M${totalRow}`).format = { fill: colors.paleBlue, font: { name: "Arial", size: 10, bold: true, color: colors.navy } };
summary.tables.add(`A6:M${products.length + 6}`, true, "ResumenProductoTable");

summary.mergeCells(`A${totalRow + 3}:M${totalRow + 3}`);
summary.getRange(`A${totalRow + 3}`).values = [["Método reproducible: para cada producto, el stock objetivo es el mayor valor entre 10 × su venta diaria promedio, 2 × su mayor venta diaria y 1 unidad. Si el stock inicial es igual o menor que 3 × el promedio diario, se registra una entrada simulada hasta el objetivo."]];
summary.getRange(`A${totalRow + 3}:M${totalRow + 3}`).format = { fill: colors.paleGray, font: { name: "Arial", size: 10, color: colors.note }, wrapText: true, verticalAlignment: "center" };
summary.getRange(`A${totalRow + 3}:M${totalRow + 3}`).format.rowHeight = 36;
summary.mergeCells(`A${totalRow + 5}:M${totalRow + 5}`);
summary.getRange(`A${totalRow + 5}`).values = [["La simulación repone antes de que el producto llegue a cero, por eso DQS y TQS resultan 0. Esto no afirma que la tienda real no tuvo quiebres; solo describe este escenario simulado."]];
summary.getRange(`A${totalRow + 5}:M${totalRow + 5}`).format = { fill: colors.green, font: { name: "Arial", size: 10, color: colors.navy, italic: true }, wrapText: true, verticalAlignment: "center" };
summary.getRange(`A${totalRow + 5}:M${totalRow + 5}`).format.rowHeight = 28;
summary.freezePanes.freezeRows(6);
summary.freezePanes.freezeColumns(4);
const summaryWidths = [43, 18, 18, 15, 17, 16, 17, 16, 17, 12, 15, 12, 12];
summaryWidths.forEach((width, index) => summary.getRangeByIndexes(0, index, totalRow + 6, 1).format.columnWidth = width);
summary.getRange(`A6:M${totalRow}`).format.rowHeight = 22;

const lastDailyRow = inventoryRows.length + 6;
daily.getRange(`A1:O${lastDailyRow}`).format.font = bodyFont;
daily.getRange(`A1:O${lastDailyRow}`).format.verticalAlignment = "center";
daily.getRange("A2").values = [["Inventario diario por producto y marca"]];
daily.getRange("A2").format.font = titleFont;
daily.mergeCells("A3:O3");
daily.getRange("A3").values = [["Cada fila equivale a un producto vendido en un día reportado. Las ventas, productos y marcas provienen de los reportes y del catálogo; inventario y quiebres son una simulación reproducible."]];
daily.getRange("A3:O3").format = { font: { name: "Arial", size: 10, color: colors.note, italic: true }, verticalAlignment: "center" };
daily.mergeCells("A4:O4");
daily.getRange("A4").values = [["Fuente de ventas y marcas: 168 reportes de ventas incorporados al sistema y catálogo maestro. Los días sin venta se consideran dentro del cálculo de stock, pero no se repiten como filas vacías en esta hoja."]];
daily.getRange("A4:O4").format = { fill: colors.amber, font: { name: "Arial", size: 10, bold: true, color: "#7F6000" }, wrapText: true, verticalAlignment: "center" };
daily.getRange("A4:O4").format.rowHeight = 28;

daily.getRange("A6:O6").values = [[
  "Fecha real",
  "Producto vendido",
  "Marca",
  "Familia",
  "Categoría",
  "Ventas CD reales",
  "Stock inicial SI sim.",
  "Entradas EN sim.",
  "Stock disponible sim.",
  "Stock final SF sim.",
  "DQS sim. (0/1)",
  "Punto reposición sim.",
  "Stock objetivo sim.",
  "Origen producto/ventas",
  "Origen inventario",
]];
daily.getRange("A6:O6").format = { fill: colors.blue, font: headerFont, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
daily.getRangeByIndexes(6, 0, inventoryRows.length, 15).values = inventoryRows;
daily.getRange(`A7:A${lastDailyRow}`).format.numberFormat = "dd/mm/yyyy";
daily.getRange(`F7:M${lastDailyRow}`).format.numberFormat = "#,##0";
daily.getRange(`F7:F${lastDailyRow}`).format.fill = colors.paleBlue;
daily.getRange(`G7:M${lastDailyRow}`).format.fill = colors.amber;
daily.getRange(`N7:N${lastDailyRow}`).format.fill = colors.paleBlue;
daily.getRange(`O7:O${lastDailyRow}`).format.fill = colors.amber;
daily.getRange(`A6:O${lastDailyRow}`).format.borders = { preset: "outside", style: "thin", color: colors.border };
daily.getRange(`A7:O${lastDailyRow}`).format.rowHeight = 18;
daily.tables.add(`A6:O${lastDailyRow}`, true, "InventarioProductoDiarioTable");
daily.freezePanes.freezeRows(6);
daily.freezePanes.freezeColumns(3);
const dailyWidths = [13, 43, 18, 18, 15, 16, 17, 16, 18, 17, 14, 19, 18, 22, 15];
dailyWidths.forEach((width, index) => daily.getRangeByIndexes(0, index, lastDailyRow, 1).format.columnWidth = width);

workbook.recalculate();
const summaryCheck = await workbook.inspect({ kind: "table", range: `Resumen por producto!A2:M${totalRow + 5}`, include: "values,formulas", tableMaxRows: 15, tableMaxCols: 13 });
console.log(summaryCheck.ndjson);
const dailyTopCheck = await workbook.inspect({ kind: "table", range: "Inventario diario!A1:O12", include: "values", tableMaxRows: 12, tableMaxCols: 15 });
console.log(dailyTopCheck.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" });
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const summaryPreview = await workbook.render({ sheetName: "Resumen por producto", range: "A1:M20", scale: 1.05, format: "png" });
await fs.writeFile(`${outputDir}/inventario_producto_resumen_preview.png`, new Uint8Array(await summaryPreview.arrayBuffer()));
const dailyTopPreview = await workbook.render({ sheetName: "Inventario diario", range: "A1:O16", scale: 0.95, format: "png" });
await fs.writeFile(`${outputDir}/inventario_producto_top_preview.png`, new Uint8Array(await dailyTopPreview.arrayBuffer()));
const dailyTailPreview = await workbook.render({ sheetName: "Inventario diario", range: `A${lastDailyRow - 15}:O${lastDailyRow}`, scale: 0.95, format: "png" });
await fs.writeFile(`${outputDir}/inventario_producto_tail_preview.png`, new Uint8Array(await dailyTailPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
