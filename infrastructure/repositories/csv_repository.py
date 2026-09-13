import os
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List

import numpy as np
import pandas as pd

from domain.interfaces.repositories import DemandRepository
from domain.entities.demand import Demand
from config.settings import HISTORIAL_FILE

class CSVDemandRepository(DemandRepository):
    REQUIRED_COLUMNS = ["date", "category", "quantity"]

    def __init__(self, file_path: Path = HISTORIAL_FILE):
        self.file_path = file_path

    def save_demands(self, demands: List[Demand]) -> dict[str, int]:
        """Guarda datos de forma atómica y distingue altas, cambios y repetidos.

        La clave funcional del historial es ``fecha + producto``. Una carga
        idéntica es idempotente; una corrección del mismo producto y fecha
        actualiza solo esa fila, sin inflar los totales.
        """
        df_new = self._normalise_dataframe(
            pd.DataFrame(
                [
                    {"date": demand.date, "category": demand.category, "quantity": demand.quantity}
                    for demand in demands
                ]
            )
        )
        if df_new.empty:
            return {"added": 0, "updated": 0, "unchanged": 0}

        df_existing = self._read_dataframe()
        keys = ["date", "category"]
        existing_lookup = df_existing.set_index(keys)["quantity"].to_dict()

        added = updated = unchanged = 0
        for row in df_new.itertuples(index=False):
            key = (row.date, row.category)
            previous = existing_lookup.get(key)
            if previous is None:
                added += 1
            elif np.isclose(float(previous), float(row.quantity)):
                unchanged += 1
            else:
                updated += 1

        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        df_combined = df_combined.drop_duplicates(subset=keys, keep="last")
        df_combined = df_combined.sort_values(keys, kind="stable").reset_index(drop=True)
        self._atomic_write(df_combined)

        return {"added": added, "updated": updated, "unchanged": unchanged}

    def get_all_demands(self) -> List[Demand]:
        df = self._read_dataframe()
        return [
            Demand(
                date=datetime.strptime(str(row["date"]), "%Y-%m-%d"),
                category=row["category"],
                quantity=float(row["quantity"]),
            )
            for _, row in df.iterrows()
        ]

    def get_demands_by_date_range(self, start_date, end_date) -> List[Demand]:
        all_demands = self.get_all_demands()
        return [
            d for d in all_demands
            if start_date <= d.date <= end_date
        ]

    def get_last_updated_at(self) -> datetime | None:
        """Devuelve la última modificación conocida del historial local."""
        if not self.file_path.exists() or self.file_path.stat().st_size == 0:
            return None
        return datetime.fromtimestamp(self.file_path.stat().st_mtime).astimezone()

    def _read_dataframe(self) -> pd.DataFrame:
        if not self.file_path.exists() or self.file_path.stat().st_size == 0:
            return pd.DataFrame(columns=self.REQUIRED_COLUMNS)
        try:
            dataframe = pd.read_csv(self.file_path)
        except pd.errors.EmptyDataError:
            return pd.DataFrame(columns=self.REQUIRED_COLUMNS)
        except (OSError, UnicodeDecodeError) as exc:
            raise ValueError(f"No se pudo leer el historial de demanda: {exc}") from exc
        return self._normalise_dataframe(dataframe)

    def _normalise_dataframe(self, dataframe: pd.DataFrame) -> pd.DataFrame:
        missing_columns = set(self.REQUIRED_COLUMNS) - set(dataframe.columns)
        if missing_columns:
            columns = ", ".join(sorted(missing_columns))
            raise ValueError(f"El historial no tiene las columnas requeridas: {columns}.")

        normalised = dataframe.loc[:, self.REQUIRED_COLUMNS].copy()
        try:
            normalised["date"] = pd.to_datetime(normalised["date"], errors="raise").dt.strftime("%Y-%m-%d")
            normalised["quantity"] = pd.to_numeric(normalised["quantity"], errors="raise")
        except (TypeError, ValueError) as exc:
            raise ValueError("El historial contiene fechas o cantidades inválidas.") from exc

        normalised["category"] = normalised["category"].astype(str).str.strip()
        if normalised["category"].eq("").any() or normalised["quantity"].isna().any():
            raise ValueError("El historial contiene productos o cantidades vacías.")
        if (normalised["quantity"] < 0).any():
            raise ValueError("El historial contiene cantidades negativas.")

        # Si el archivo de origen repite una fila, se conserva su última
        # versión. Dentro de una carga, IngestService ya suma el producto/día.
        return normalised.drop_duplicates(subset=["date", "category"], keep="last").reset_index(drop=True)

    def _atomic_write(self, dataframe: pd.DataFrame) -> None:
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                newline="",
                suffix=".tmp",
                prefix=f".{self.file_path.stem}.",
                dir=self.file_path.parent,
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)
                dataframe.to_csv(temporary_file, index=False)
            os.replace(temporary_path, self.file_path)
        except PermissionError as exc:
            raise ValueError(
                "No se puede actualizar el historial porque 'historial_demanda.csv' está abierto "
                "en otra aplicación. Ciérralo en Excel y vuelve a procesar el reporte."
            ) from exc
        finally:
            if temporary_path is not None and temporary_path.exists():
                temporary_path.unlink(missing_ok=True)
