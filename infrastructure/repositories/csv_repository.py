import pandas as pd
from typing import List
from datetime import datetime
from pathlib import Path

from domain.interfaces.repositories import DemandRepository
from domain.entities.demand import Demand
from config.settings import HISTORIAL_FILE

class CSVDemandRepository(DemandRepository):
    def __init__(self, file_path: Path = HISTORIAL_FILE):
        self.file_path = file_path

    def save_demands(self, demands: List[Demand]) -> None:
        # Convertir a DataFrame
        data = [
            {'date': d.date, 'category': d.category, 'quantity': d.quantity}
            for d in demands
        ]
        df_new = pd.DataFrame(data)

        # Cargar histórico existente SOLO si el archivo existe y NO está vacío
        if self.file_path.exists() and self.file_path.stat().st_size > 0:
            try:
                df_existing = pd.read_csv(self.file_path, parse_dates=['date'])
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined = df_combined.drop_duplicates(subset=['date', 'category'], keep='last')
            except pd.errors.EmptyDataError:
                # Archivo vacío, ignoramos y usamos solo los nuevos
                df_combined = df_new
        else:
            df_combined = df_new

        # Guardar. En Windows, Excel bloquea el archivo mientras está abierto;
        # devolvemos una instrucción accionable para la interfaz web.
        try:
            df_combined.to_csv(self.file_path, index=False)
        except PermissionError as exc:
            raise ValueError(
                "No se puede actualizar el historial porque 'historial_demanda.csv' está abierto "
                "en otra aplicación. Ciérralo en Excel y vuelve a procesar el reporte."
            ) from exc

    def get_all_demands(self) -> List[Demand]:
        if not self.file_path.exists() or self.file_path.stat().st_size == 0:
            return []
        try:
            df = pd.read_csv(self.file_path, parse_dates=['date'])
            return [
                Demand(date=row['date'], category=row['category'], quantity=row['quantity'])
                for _, row in df.iterrows()
            ]
        except pd.errors.EmptyDataError:
            return []

    def get_demands_by_date_range(self, start_date, end_date) -> List[Demand]:
        all_demands = self.get_all_demands()
        return [
            d for d in all_demands
            if start_date <= d.date <= end_date
        ]
