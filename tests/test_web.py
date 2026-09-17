from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch
import unittest

from interfaces import web
from domain.entities.demand import Demand
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.repositories.csv_repository import CSVDemandRepository


class WebTests(unittest.TestCase):
    def test_parse_cutoff_date_accepts_iso_dates_and_rejects_invalid_values(self):
        self.assertEqual(web._parse_cutoff_date("2026-06-15"), date(2026, 6, 15))
        self.assertIsNone(web._parse_cutoff_date(""))
        with self.assertRaisesRegex(ValueError, "formato AAAA-MM-DD"):
            web._parse_cutoff_date("15/06/2026")

    def test_historical_cutoff_is_sent_to_training_and_forecast(self):
        class TrainServiceStub:
            received_cutoff = None

            def __init__(self, _repository):
                self.last_training_summary = {"trained_categories": 1}

            def run(self, cutoff_date=None):
                TrainServiceStub.received_cutoff = cutoff_date
                return {"wape": 10.0, "rmse": 2.0, "mae": 1.0}

        class PredictServiceStub:
            received_cutoff = None

            def __init__(self, _repository):
                self.predictor = SimpleNamespace(metadata={"validation_by_category": {}})

            def predict_future(self, days, cutoff_date=None):
                PredictServiceStub.received_cutoff = cutoff_date
                return {
                    "ABARROTES": [
                        Demand(datetime(2026, 6, 16), "ABARROTES", 4.0)
                        for _ in range(days)
                    ]
                }

        with TemporaryDirectory() as folder:
            model_path = Path(folder) / "models.pkl"
            model_path.touch()
            with patch.object(web, "TrainService", TrainServiceStub), patch.object(
                web, "PredictService", PredictServiceStub
            ), patch.object(web, "MODELS_FILE", model_path), patch.object(
                web, "_load_category_products", return_value={"ABARROTES": []}
            ):
                client = web.create_app({"TESTING": True}).test_client()
                trained = client.post("/api/train", json={"cutoff_date": "2026-06-15"})
                forecast = client.post("/api/forecast", json={"days": 2, "cutoff_date": "2026-06-15"})

            self.assertEqual(trained.status_code, 200)
            self.assertEqual(trained.get_json()["cutoff_date"], "2026-06-15")
            self.assertEqual(TrainServiceStub.received_cutoff, date(2026, 6, 15))
            self.assertEqual(forecast.status_code, 200)
            self.assertEqual(forecast.get_json()["cutoff_date"], "2026-06-15")
            self.assertEqual(PredictServiceStub.received_cutoff, date(2026, 6, 15))

    def test_dashboard_category_percentages_cover_the_full_history_and_recalculate(self):
        class CatalogStub:
            def get_summary(self):
                return {"products": 2, "families": 1, "categories": 2}

            def get_categories(self):
                return ["ABARROTES", "BEBIDAS"]

        with TemporaryDirectory() as folder:
            repository = CSVDemandRepository(Path(folder) / "history.csv")
            repository.save_demands([
                Demand(datetime(2026, 1, 1), "ABARROTES", 10),
                Demand(datetime(2026, 1, 1), "BEBIDAS", 30),
            ])

            first_dashboard = web._build_dashboard(repository, CatalogStub())
            self.assertEqual(first_dashboard["summary"]["first_sale_date"], "2026-01-01")
            self.assertEqual(first_dashboard["summary"]["last_sale_date"], "2026-01-01")
            self.assertEqual(first_dashboard["categories"], [
                {"name": "BEBIDAS", "quantity": 30.0, "percentage": 75.0},
                {"name": "ABARROTES", "quantity": 10.0, "percentage": 25.0},
            ])

            repository.save_demands([Demand(datetime(2026, 1, 2), "ABARROTES", 30)])
            updated_dashboard = web._build_dashboard(repository, CatalogStub())
            self.assertEqual(updated_dashboard["summary"]["first_sale_date"], "2026-01-01")
            self.assertEqual(updated_dashboard["summary"]["last_sale_date"], "2026-01-02")
            self.assertEqual(updated_dashboard["categories"], [
                {"name": "ABARROTES", "quantity": 40.0, "percentage": 57.14},
                {"name": "BEBIDAS", "quantity": 30.0, "percentage": 42.86},
            ])

    def test_category_product_ranking_includes_historical_units(self):
        ranking = web._load_category_products()

        self.assertEqual(ranking["ABARROTES"][0]["rank"], 1)
        self.assertTrue(ranking["ABARROTES"][0]["name"])
        self.assertGreater(ranking["ABARROTES"][0]["historical_quantity"], 0)
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

                report_path = Path(__file__).resolve().parents[1] / "data" / "reporte_ventas_2026-09-13.xlsx"
                expected_matched_products = len({sale.product_name for sale in ExcelReader().read_sales(str(report_path))})
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
                self.assertEqual(payload["catalog_summary"]["matched_products"], expected_matched_products)
                self.assertEqual(payload["catalog_summary"]["unmatched_products"], [])
                self.assertEqual(payload["report_dates"], ["2026-09-13"])
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
