import fs from "node:fs/promises";

const inputPath = "C:/taully_demand_forecast/.tmp_inventory_demo/product_sales_from_reports.json";
const input = JSON.parse(await fs.readFile(inputPath, "utf8"));
const { metadata, records: sales } = input;
const dates = Array.from({ length: metadata.date_count }, (_, index) => {
  const [year, month, day] = metadata.start_date.split("-").map(Number);
  return new Date(year, month - 1, day + index).toISOString().slice(0, 10);
});

const products = new Map();
const salesByDateProduct = new Map();
for (const sale of sales) {
  products.set(sale.product, sale);
  salesByDateProduct.set(`${sale.date}|${sale.product}`, sale.quantity);
}

const metrics = new Map();
function categoryMetrics(category) {
  if (!metrics.has(category)) {
    metrics.set(category, { si: 0, entries: 0, sf: 0, sales: 0, stockoutEvents: 0, stockoutDates: new Set() });
  }
  return metrics.get(category);
}

for (const product of products.values()) {
  const dailySales = dates.map((date) => salesByDateProduct.get(`${date}|${product.product}`) ?? 0);
  const average = dailySales.reduce((sum, quantity) => sum + quantity, 0) / dates.length;
  const highestSale = Math.max(...dailySales);
  const stockTarget = Math.max(Math.ceil(average * 10), highestSale * 2, 1);
  const reorderPoint = Math.max(Math.ceil(average * 3), 1);
  const category = categoryMetrics(product.category);
  let stockInitial = stockTarget;
  category.si += stockInitial;

  for (let index = 0; index < dates.length; index += 1) {
    const quantity = dailySales[index];
    const entries = stockInitial <= reorderPoint ? stockTarget - stockInitial : 0;
    const available = stockInitial + entries;
    if (quantity > available) {
      category.stockoutEvents += 1;
      category.stockoutDates.add(dates[index]);
    }
    category.entries += entries;
    category.sales += quantity;
    stockInitial = Math.max(0, available - quantity);
  }
  category.sf += stockInitial;
}

const result = [...metrics.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([category, value]) => ({
  category,
  stock_initial: value.si,
  entries: value.entries,
  stock_final: value.sf,
  demand: value.sales,
  inventory_exit_index_pct: Number((value.sales / (value.si + value.entries) * 100).toFixed(2)),
  stockout_product_day_events: value.stockoutEvents,
  stockout_category_days: value.stockoutDates.size,
  stockout_rate_category_days_pct: Number((value.stockoutDates.size / dates.length * 100).toFixed(2)),
}));
console.log(JSON.stringify({ period: `${metadata.start_date} to ${metadata.end_date}`, days: dates.length, result }, null, 2));
