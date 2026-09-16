from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import unittest

from application.services.predict_service import PredictService
from application.services.train_service import TrainService
from domain.entities.demand import Demand
from infrastructure.repositories.csv_repository import CSVDemandRepository


class ForecastingTests(unittest.TestCase):
    def _repository_with_varied_data(self, folder: str) -> CSVDemandRepository:
        repository = CSVDemandRepository(Path(folder) / "history.csv")
        start = datetime(2026, 1, 1)
        demands = []
        for day in range(70):
            date = start + timedelta(days=day)
            demands.extend(
                [
                    Demand(date, "PRODUCTO A", 5 + (day % 7)),
                    Demand(date, "PRODUCTO B", 12 + (day % 7)),
                ]
            )
        repository.save_demands(demands)
        return repository

    def test_train_predict_and_detect_stale_model(self):
        with TemporaryDirectory() as folder:
            repository = self._repository_with_varied_data(folder)
            model_path = Path(folder) / "models.pkl"
            with patch("application.services.train_service.MODELS_FILE", model_path), patch(
                "application.services.predict_service.MODELS_FILE", model_path
            ):
                metrics = TrainService(repository).run()
                self.assertTrue(model_path.exists())
                self.assertEqual(set(metrics), {"mae", "rmse", "wape"})

                predictions = PredictService(repository).predict_future(3)
                self.assertEqual(set(predictions), {"PRODUCTO A", "PRODUCTO B"})
                self.assertTrue(all(len(rows) == 3 for rows in predictions.values()))
                self.assertTrue(all(row.quantity >= 0 for rows in predictions.values() for row in rows))

                repository.save_demands([Demand(datetime(2026, 2, 5), "PRODUCTO A", 9)])
                with self.assertRaisesRegex(ValueError, "historial cambió"):
                    PredictService(repository).predict_future(1)
