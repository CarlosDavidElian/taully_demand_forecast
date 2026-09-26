from __future__ import annotations

from collections import defaultdict
from datetime import datetime
import json
from pathlib import Path
import sys

project = Path("C:/taully_demand_forecast")
sys.path.insert(0, str(project))
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository


reports = sorted((project / "data").glob("reporte_ventas_*.xlsx"))
catalog = ExcelCatalogRepository(project / "data" / "catalogo_maestro.xlsx")
reader = ExcelReader()

daily_sales: dict[tuple[str, str], float] = defaultdict(float)
product_details: dict[str, dict[str, str]] = {}
unmatched: dict[str, float] = defaultdict(float)
report_dates: set[str] = set()
source_units = 0.0
categorized_units = 0.0

for report_path in reports:
    sales = reader.read_sales(str(report_path))
    if not sales:
        raise ValueError(f"El reporte {report_path.name} no contiene ventas válidas.")

    dates_in_report = {sale.date.strftime("%Y-%m-%d") for sale in sales}
    if len(dates_in_report) != 1:
        raise ValueError(f"El reporte {report_path.name} tiene más de una fecha.")
    report_date = dates_in_report.pop()
    if report_date in report_dates:
        raise ValueError(f"Hay más de un reporte para {report_date}.")
    report_dates.add(report_date)

    for sale in sales:
        source_units += float(sale.quantity)
        catalog_product = catalog.get_product(sale.product_name)
        if catalog_product is None:
            unmatched[sale.product_name] += float(sale.quantity)
            continue

        product_name = catalog_product.product_name
        daily_sales[(report_date, product_name)] += float(sale.quantity)
        product_details[product_name] = {
            "product": product_name,
            "brand": catalog_product.brand or "Sin marca registrada",
            "family": catalog_product.family,
            "category": catalog_product.category,
        }
        categorized_units += float(sale.quantity)

records = []
for (report_date, product_name), quantity in sorted(daily_sales.items()):
    records.append({"date": report_date, **product_details[product_name], "quantity": quantity})

output = {
    "metadata": {
        "report_count": len(reports),
        "date_count": len(report_dates),
        "start_date": min(report_dates),
        "end_date": max(report_dates),
        "source_units": source_units,
        "categorized_units": categorized_units,
        "unmatched_units": source_units - categorized_units,
        "unmatched_products": len(unmatched),
        "product_count": len(product_details),
        "sales_record_count": len(records),
    },
    "records": records,
    "unmatched": [
        {"product": product, "quantity": quantity}
        for product, quantity in sorted(unmatched.items())
    ],
}

output_path = project / ".tmp_inventory_demo" / "product_sales_from_reports.json"
output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")

print(json.dumps(output["metadata"], ensure_ascii=False, indent=2))
