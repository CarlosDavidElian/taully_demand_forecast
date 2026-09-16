from __future__ import annotations

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import unittest

from interfaces import web
from infrastructure.repositories.csv_repository import CSVDemandRepository


class WebTests(unittest.TestCase):
    def test_category_product_ranking_includes_historical_units(self):
        ranking = web._load_category_products()

        self.assertEqual(ranking["ABARROTES"][0]["rank"], 1)
        self.assertEqual(ranking["ABARROTES"][0]["name"], "MACA NEGRA CON")
        self.assertEqual(ranking["ABARROTES"][0]["historical_quantity"], 370.0)
        self.assertGreater(ranking["ABARROTES"][0]["historical_share"], 0)
        self.assertEqual(ranking["ABARROTES"][1]["rank"], 2)
        self.assertGreater(ranking["ABARROTES"][0]["historical_quantity"], ranking["ABARROTES"][1]["historical_quantity"])
        self.assertGreater(len(ranking["ABARROTES"]), 5)

    def test_product_forecast_allocations_reconcile_to_category_total(self):
        product_forecast = web._build_product_forecasts(
            {"ABARROTES": [{"date": "2026-09-14", "quantity": 100.0}]},
            {
                "ABARROTES": [
                    {"name": "PRODUCTO A", "historical_share": 0.6},
                    {"name": "PRODUCTO B", "historical_share": 0.4},
                ]
            },
        )

        detail = product_forecast["ABARROTES"][0]
        self.assertEqual(detail["category_quantity"], 100.0)
        self.assertEqual(detail["products"], [
            {"name": "PRODUCTO A", "quantity": 60.0, "historical_share": 60.0},
            {"name": "PRODUCTO B", "quantity": 40.0, "historical_share": 40.0},
        ])
        self.assertEqual(sum(item["quantity"] for item in detail["products"]), detail["category_quantity"])

    def test_dashboard_and_upload_report(self):
        with TemporaryDirectory() as folder:
            repository = CSVDemandRepository(Path(folder) / "history.csv")
            with patch.object(web, "CSVDemandRepository", return_value=repository):
                app = web.create_app({"TESTING": True})
                client = app.test_client()

                dashboard = client.get("/api/dashboard")
                self.assertEqual(dashboard.status_code, 200)
                self.assertEqual(dashboard.headers["Cache-Control"], "no-store")
                dashboard_payload = dashboard.get_json()
                self.assertEqual(dashboard_payload["summary"]["records"], 0)
                self.assertEqual(dashboard_payload["summary"]["categories"], 0)
                self.assertIsNone(dashboard_payload["summary"]["first_sale_date"])
                self.assertGreater(dashboard_payload["catalog"]["products"], 0)

                catalog = client.get("/api/catalog")
                self.assertEqual(catalog.status_code, 200)
                self.assertEqual(catalog.get_json()["summary"]["products"], dashboard_payload["catalog"]["products"])

                favicon = client.get("/favicon.ico")
                try:
                    self.assertEqual(favicon.status_code, 200)
                    self.assertEqual(favicon.mimetype, "image/vnd.microsoft.icon")
                finally:
                    favicon.close()

                report_path = Path(__file__).resolve().parents[1] / "data" / "Reporte_Taully_2026-09-13.xlsx"
                with report_path.open("rb") as report:
                    response = client.post(
                        "/api/reports",
                        data={"report": (report, report_path.name)},
                        content_type="multipart/form-data",
                    )
                payload = response.get_json()
                self.assertEqual(response.status_code, 200)
                expected_categories = len(web.CatalogService(web.ExcelCatalogRepository()).get_categories())
                self.assertEqual(payload["save_summary"], {"added": expected_categories, "updated": 0, "unchanged": 0})
                self.assertEqual(payload["catalog_summary"]["matched_products"], 14)
                self.assertEqual(payload["catalog_summary"]["unmatched_products"], [])
                self.assertEqual(payload["dashboard"]["summary"]["last_sale_date"], "2026-09-13")
                self.assertEqual(payload["dashboard"]["summary"]["first_sale_date"], "2026-09-13")
                self.assertEqual(payload["dashboard"]["summary"]["categories"], expected_categories)

    def test_catalog_upload_replaces_only_a_validated_file(self):
        with TemporaryDirectory() as folder:
            repository = CSVDemandRepository(Path(folder) / "history.csv")
            target_catalog = Path(folder) / "catalogo_maestro.xlsx"
            with patch.object(web, "CSVDemandRepository", return_value=repository):
                app = web.create_app({"TESTING": True, "CATALOG_FILE": target_catalog})
                client = app.test_client()

                invalid = client.post(
                    "/api/catalog",
                    data={"catalog": (BytesIO(b"esto no es un archivo Excel"), "catalogo.xlsx")},
                    content_type="multipart/form-data",
                )
                self.assertEqual(invalid.status_code, 400)
                self.assertFalse(target_catalog.exists())

                source_catalog = Path(__file__).resolve().parents[1] / "data" / "catalogo_maestro.xlsx"
                with source_catalog.open("rb") as catalog:
                    response = client.post(
                        "/api/catalog",
                        data={"catalog": (catalog, source_catalog.name)},
                        content_type="multipart/form-data",
                    )

                self.assertEqual(response.status_code, 200)
                self.assertTrue(target_catalog.exists())
                self.assertGreater(response.get_json()["catalog"]["products"], 0)
