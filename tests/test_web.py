from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import unittest

from interfaces import web
from infrastructure.repositories.csv_repository import CSVDemandRepository


class WebTests(unittest.TestCase):
    def test_dashboard_and_upload_report(self):
        with TemporaryDirectory() as folder:
            repository = CSVDemandRepository(Path(folder) / "history.csv")
            with patch.object(web, "CSVDemandRepository", return_value=repository):
                app = web.create_app({"TESTING": True})
                client = app.test_client()

                dashboard = client.get("/api/dashboard")
                self.assertEqual(dashboard.status_code, 200)
                self.assertEqual(dashboard.headers["Cache-Control"], "no-store")
                self.assertEqual(dashboard.get_json()["summary"]["records"], 0)

                favicon = client.get("/favicon.ico")
                try:
                    self.assertEqual(favicon.status_code, 200)
                    self.assertEqual(favicon.mimetype, "image/vnd.microsoft.icon")
                finally:
                    favicon.close()

                report_path = Path(__file__).resolve().parents[1] / "data" / "reporte_20260801.xlsx"
                with report_path.open("rb") as report:
                    response = client.post(
                        "/api/reports",
                        data={"report": (report, report_path.name)},
                        content_type="multipart/form-data",
                    )
                payload = response.get_json()
                self.assertEqual(response.status_code, 200)
                self.assertEqual(payload["save_summary"], {"added": 9, "updated": 0, "unchanged": 0})
                self.assertEqual(payload["dashboard"]["summary"]["last_sale_date"], "2026-08-01")
                self.assertEqual(payload["dashboard"]["summary"]["products"], 9)
