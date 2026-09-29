from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
import json
import unittest
from unittest.mock import patch

from application.services.posttest_evaluation_service import PosttestEvaluationService
import application.services.predict_service as predict_service_module
import application.services.train_service as train_service_module
from domain.entities.demand import Demand
from infrastructure.repositories.csv_repository import CSVDemandRepository


class PosttestEvaluationServiceTests(unittest.TestCase):
    def _repository(self, folder: str, days: int = 70) -> CSVDemandRepository:
        repository = CSVDemandRepository(Path(folder) / "history.csv")
        start = datetime(2026, 1, 1)
        demands: list[Demand] = []
        for day in range(days):
            current_date = start + timedelta(days=day)
            demands.extend(
                [
                    Demand(current_date, "ABARROTES", 20 + (day % 7)),
                    Demand(current_date, "BEBIDAS", 35 + ((day * 2) % 7)),
                ]
            )
        repository.save_demands(demands)
        return repository

    def test_runs_a_serialisable_posttest_without_writing_the_user_model(self):
        with TemporaryDirectory() as folder:
            repository = self._repository(folder)
            user_model_path = Path(folder) / "user_models_by_category.pkl"
            user_model_path.write_bytes(b"modelo normal del usuario")
            with patch.object(train_service_module, "MODELS_FILE", user_model_path), patch.object(
                predict_service_module, "MODELS_FILE", user_model_path
            ):
                result = PosttestEvaluationService(repository).run(
                    cutoff_date=datetime(2026, 3, 1).date(),
                    horizon_days=7,
                )

            self.assertEqual(result["cutoff_date"], "2026-03-01")
            self.assertEqual(result["period"], {"start_date": "2026-03-02", "end_date": "2026-03-08"})
            self.assertEqual(result["categories_evaluated"], ["ABARROTES", "BEBIDAS"])
            self.assertEqual(result["observation_count"], 14)
            self.assertTrue(result["training"]["temporary_model_artifact"])
            self.assertEqual(len(result["observations"]), 14)
            self.assertIn("mape", result["model"]["metrics"])
            self.assertIn("wape", result["baseline"]["metrics"])
            self.assertEqual(len(result["methodology"]["history_fingerprint"]), 64)
            self.assertEqual(len(result["methodology"]["training_history_fingerprint"]), 64)
            self.assertEqual(result["methodology"]["history_records"], 140)
            self.assertEqual(result["methodology"]["training_history_records"], 120)
            self.assertEqual(set(result["comparison"]["winner_by_metric"]), {"mae", "rmse", "mape", "wape"})
            self.assertEqual(user_model_path.read_bytes(), b"modelo normal del usuario")

            # JSON es lo que consumirán la pantalla y el exportador.
            json.dumps(result, allow_nan=False)

    def test_global_metrics_are_calculated_from_the_posttest_observations(self):
        with TemporaryDirectory() as folder:
            repository = self._repository(folder)
            result = PosttestEvaluationService(repository).evaluate(
                cutoff_date=datetime(2026, 3, 1),
                horizon_days=7,
            )

            observations = result["observations"]
            actual_total = sum(row["actual_demand"] for row in observations)
            model_total = sum(row["model_prediction"] for row in observations)
            absolute_error_total = sum(
                abs(row["actual_demand"] - row["model_prediction"])
                for row in observations
            )

            self.assertAlmostEqual(result["model"]["metrics"]["actual_total"], actual_total, places=5)
            self.assertAlmostEqual(result["model"]["metrics"]["prediction_total"], model_total, places=5)
            self.assertAlmostEqual(
                result["model"]["metrics"]["wape"],
                absolute_error_total / actual_total * 100,
                places=5,
            )

    def test_requires_complete_real_sales_after_the_cutoff(self):
        with TemporaryDirectory() as folder:
            repository = self._repository(folder, days=65)
            service = PosttestEvaluationService(repository)

            with self.assertRaisesRegex(ValueError, "No hay ventas reales completas"):
                service.run(cutoff_date=datetime(2026, 3, 1).date(), horizon_days=7)
