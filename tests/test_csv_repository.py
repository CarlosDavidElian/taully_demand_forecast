from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from domain.entities.demand import Demand
from infrastructure.repositories.csv_repository import CSVDemandRepository


class CSVDemandRepositoryTests(unittest.TestCase):
    def test_reports_added_updated_and_unchanged_rows(self):
        with TemporaryDirectory() as folder:
            repository = CSVDemandRepository(Path(folder) / "history.csv")
            demand = Demand(datetime(2026, 9, 11), "PRODUCTO A", 5)

            self.assertEqual(repository.save_demands([demand]), {"added": 1, "updated": 0, "unchanged": 0})
            self.assertEqual(repository.save_demands([demand]), {"added": 0, "updated": 0, "unchanged": 1})
            self.assertEqual(
                repository.save_demands([Demand(datetime(2026, 9, 11), "PRODUCTO A", 8)]),
                {"added": 0, "updated": 1, "unchanged": 0},
            )

            saved = repository.get_all_demands()
            self.assertEqual(len(saved), 1)
            self.assertEqual(saved[0].quantity, 8)

