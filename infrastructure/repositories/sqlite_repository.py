import sqlite3
from typing import List
from datetime import datetime
from pathlib import Path

from domain.interfaces.repositories import DemandRepository
from domain.entities.demand import Demand
from config.settings import DATA_DIR

class SQLiteDemandRepository(DemandRepository):
    def __init__(self, db_path: Path = DATA_DIR / "historial.db"):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS demands (
                    date TEXT NOT NULL,
                    category TEXT NOT NULL,
                    quantity REAL NOT NULL,
                    PRIMARY KEY (date, category)
                )
            """)

    def save_demands(self, demands: List[Demand]) -> dict[str, int]:
        added = updated = unchanged = 0
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            for d in demands:
                previous = cursor.execute(
                    "SELECT quantity FROM demands WHERE date = ? AND category = ?",
                    (d.date.isoformat(), d.category),
                ).fetchone()
                if previous is None:
                    added += 1
                elif float(previous[0]) == float(d.quantity):
                    unchanged += 1
                else:
                    updated += 1
                cursor.execute(
                    "INSERT OR REPLACE INTO demands (date, category, quantity) VALUES (?, ?, ?)",
                    (d.date.isoformat(), d.category, d.quantity)
                )
            conn.commit()
        return {"added": added, "updated": updated, "unchanged": unchanged}

    def get_all_demands(self) -> List[Demand]:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute("SELECT date, category, quantity FROM demands ORDER BY date")
            rows = cursor.fetchall()
            return [
                Demand(date=datetime.fromisoformat(row[0]), category=row[1], quantity=row[2])
                for row in rows
            ]

    def get_demands_by_date_range(self, start_date, end_date) -> List[Demand]:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute(
                "SELECT date, category, quantity FROM demands WHERE date BETWEEN ? AND ? ORDER BY date",
                (start_date.isoformat(), end_date.isoformat())
            )
            rows = cursor.fetchall()
            return [
                Demand(date=datetime.fromisoformat(row[0]), category=row[1], quantity=row[2])
                for row in rows
            ]
