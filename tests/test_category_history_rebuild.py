from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from application.services.category_history_rebuild_service import CategoryHistoryRebuildService
from domain.entities.demand import Demand
from domain.entities.product import Product
from domain.entities.sale import Sale
from infrastructure.repositories.csv_repository import CSVDemandRepository


class FakeCatalogService:
    def __init__(self):
        self.products = {
            "PRODUCTO A": Product("PRODUCTO A", "ALIMENTOS", "CATEGORÍA A"),
            "PRODUCTO B": Product("PRODUCTO B", "BEBIDAS", "CATEGORÍA B"),
        }

    def get_all_products(self):
        return list(self.products.values())

    def get_product(self, product_name: str):
        return self.products.get(product_name)


class FakeSaleReader:
    def __init__(self, sales_by_name):
        self.sales_by_name = sales_by_name

    def read_sales(self, file_path: str):
        return self.sales_by_name[Path(file_path).name]


class CategoryHistoryRebuildTests(unittest.TestCase):
    def test_rebuilds_complete_daily_category_panel_and_keeps_audit_trail(self):
        with TemporaryDirectory() as folder:
            directory = Path(folder)
            report_one = directory / "Reporte_Taully_2026-01-01.xlsx"
            report_two = directory / "Reporte_Taully_2026-01-02.xlsx"
            report_one.touch()
            report_two.touch()

            sales_by_name = {
                report_one.name: [
                    Sale(datetime(2026, 1, 1), "PRODUCTO A", 5, 10),
                    Sale(datetime(2026, 1, 1), "SIN CATÁLOGO", 3, 6),
                ],
                report_two.name: [Sale(datetime(2026, 1, 2), "PRODUCTO B", 7, 14)],
            }
            repository = CSVDemandRepository(directory / "history.csv")
            repository.save_demands([Demand(datetime(2025, 12, 31), "LEGACY", 99)])
            service = CategoryHistoryRebuildService(
                demand_repo=repository,
                catalog_service=FakeCatalogService(),
                sale_reader=FakeSaleReader(sales_by_name),
                backup_directory=directory / "backups",
                uncategorized_sales_path=directory / "ventas_sin_categoria.csv",
                category_product_mix_path=directory / "productos_por_categoria.csv",
            )

            result = service.rebuild([report_two, report_one])

            self.assertEqual(result.reports, 2)
            self.assertEqual(result.dates, 2)
            self.assertEqual(result.categories, 2)
            self.assertEqual(result.demand_records, 4)
            self.assertEqual(result.source_units, 15)
            self.assertEqual(result.categorized_units, 12)
            self.assertEqual(result.uncategorized_units, 3)
            self.assertEqual(result.uncategorized_products, 1)
            self.assertTrue(result.backup_path.exists())
            self.assertTrue(result.uncategorized_sales_path.exists())
            self.assertTrue(result.category_product_mix_path.exists())

            rebuilt = {(row.date.strftime("%Y-%m-%d"), row.category): row.quantity for row in repository.get_all_demands()}
            self.assertEqual(
                rebuilt,
                {
                    ("2026-01-01", "CATEGORÍA A"): 5,
                    ("2026-01-01", "CATEGORÍA B"): 0,
                    ("2026-01-02", "CATEGORÍA A"): 0,
                    ("2026-01-02", "CATEGORÍA B"): 7,
                },
            )
            self.assertIn("SIN CATÁLOGO", result.uncategorized_sales_path.read_text(encoding="utf-8"))
            self.assertIn("PRODUCTO A", result.category_product_mix_path.read_text(encoding="utf-8"))
