import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = "C:/taully_demand_forecast/.tmp_inventory_demo/product_sales_from_reports.json";
const outputDir = "C:/taully_demand_forecast/outputs/01a0b651-7620-73f0-8f6a-8b68b14f7a17";
const outputPath = `${outputDir}/datos_prueba_inventario_producto_marca_2026-04-01_a_2026-09-15.xlsx`;

function parseDate(dateText) {
  const [year, month, day] = dateText.split("-").map(Number);
  return new Date(year, month - 1, day);
}

const input = JSON.parse(await fs.readFile(inputPath, "utf8"));
const { metadata, records: sales } = input;
const dates = Array.from({ length: metadata.date_count }, (_, index) => {
  const start = parseDate(metadata.start_date);
  return new Date(start.getFullYear(), start.getMonth(), start.getDate() + index);
});
const dateKeys = dates.map((date) => date.toISOString().slice(0, 10));
const products = new Map();
const salesByDateProduct = new Map();

for (const sale of sales) {
  products.set(sale.product, sale);
  salesByDateProduct.set(`${sale.date}|${sale.product}`, sale.quantity);
}

const rows = [];
for (const product of [...products.values()].sort((a, b) =>
  a.category.localeCompare(b.category) || a.brand.localeCompare(b.brand) || a.product.localeCompare(b.product),
)) {
  const dailySales = dateKeys.map((date) => salesByDateProduct.get(`${date}|${product.product}`) ?? 0);
  const average = dailySales.reduce((sum, quantity) => sum + quantity, 0) / dateKeys.length;
  const highestSale = Math.max(...dailySales);
  const stockTarget = Math.max(Math.ceil(average * 10), highestSale * 2, 1);
  const reorderPoint = Math.max(Math.ceil(average * 3), 1);
  let stockInitial = stockTarget;

  for (let index = 0; index < dates.length; index += 1) {
    const quantity = dailySales[index];
    const entries = stockInitial <= reorderPoint ? stockTarget - stockInitial : 0;
    const available = stockInitial + entries;
    const daysWithoutStock = quantity > available ? 1 : 0;
    const stockFinal = Math.max(0, available - quantity);

    if (quantity > 0) {
      rows.push([
        dates[index],
        product.product,
        product.brand,
        product.family,
        product.category,
        quantity,
        stockInitial,
        entries,
        available,
        stockFinal,
        daysWithoutStock,
        reorderPoint,
        stockTarget,
      ]);
    }
    stockInitial = stockFinal;
  }
}

const totalSales = rows.reduce((sum, row) => sum + row[5], 0);
if (totalSales !== metadata.categorized_units) {
  throw new Error(`Las ventas no cuadran: ${totalSales} vs ${metadata.categorized_units}.`);
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Datos inventario");
sheet.showGridLines = false;

const headers = [
  "Fecha",
  "Producto",
  "Marca",
  "Familia",
  "Categoría",
  "Cantidad_Vendida",
  "Stock_Inicial",
  "Entradas",
  "Stock_Disponible",
  "Stock_Final",
  "Dias_Sin_Stock",
  "Punto_Reorden",
  "Stock_Objetivo",
];
sheet.getRange("A1:M1").values = [headers];
sheet.getRange("A1:M1").format = {
  fill: "#245580",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
sheet.getRangeByIndexes(1, 0, rows.length, headers.length).values = rows;
sheet.getRange(`A2:A${rows.length + 1}`).format.numberFormat = "dd/mm/yyyy";
sheet.getRange(`F2:M${rows.length + 1}`).format.numberFormat = "#,##0";
sheet.getRange(`A1:M${rows.length + 1}`).format.font = { name: "Arial", size: 10, color: "#1F2937" };
sheet.getRange("A1:M1").format.font = { name: "Arial", size: 10, bold: true, color: "#FFFFFF" };
sheet.getRange(`A1:M${rows.length + 1}`).format.verticalAlignment = "center";
sheet.getRange(`A1:M${rows.length + 1}`).format.borders = { preset: "outside", style: "thin", color: "#D9E2F3" };
sheet.getRange(`A2:M${rows.length + 1}`).format.rowHeight = 18;
sheet.tables.add(`A1:M${rows.length + 1}`, true, "DatosPruebaInventarioTable");
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(3);

const widths = [13, 43, 18, 18, 15, 17, 15, 13, 18, 15, 17, 16, 16];
widths.forEach((width, index) => sheet.getRangeByIndexes(0, index, rows.length + 1, 1).format.columnWidth = width);

workbook.recalculate();
const check = await workbook.inspect({ kind: "table", range: "Datos inventario!A1:M12", include: "values", tableMaxRows: 12, tableMaxCols: 13 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "formula errors" });
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const topPreview = await workbook.render({ sheetName: "Datos inventario", range: "A1:M16", scale: 1.1, format: "png" });
await fs.writeFile(`${outputDir}/datos_prueba_inventario_top_preview.png`, new Uint8Array(await topPreview.arrayBuffer()));
const tailPreview = await workbook.render({ sheetName: "Datos inventario", range: `A${rows.length - 13}:M${rows.length + 1}`, scale: 1.1, format: "png" });
await fs.writeFile(`${outputDir}/datos_prueba_inventario_tail_preview.png`, new Uint8Array(await tailPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Rows: ${rows.length}; sales: ${totalSales}; saved ${outputPath}`);
