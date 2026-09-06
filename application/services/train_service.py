import pandas as pd
import numpy as np
from typing import Dict, Tuple

from domain.interfaces.repositories import DemandRepository
from infrastructure.ml.model_trainer import ModelTrainer
from config.settings import BEST_MODEL_FILE

class TrainService:
    def __init__(self, demand_repo: DemandRepository):
        self.demand_repo = demand_repo
        self.trainer = ModelTrainer()

    def run(self) -> Dict[str, float]:
        # 1. Obtener historial
        demands = self.demand_repo.get_all_demands()
        if len(demands) < 5:
            raise ValueError("Se necesitan al menos 30 días de datos históricos para entrenar")

        # 2. Convertir a DataFrame y crear features temporales
        df = pd.DataFrame([
            {'date': d.date, 'category': d.category, 'quantity': d.quantity}
            for d in demands
        ])
        df = df.sort_values('date').reset_index(drop=True)

        # 3. Preparar datos para entrenamiento (por categoría)
        # Vamos a entrenar un modelo por categoría para mayor precisión
        categories = df['category'].unique()
        best_overall_metrics = {}

        for cat in categories:
            df_cat = df[df['category'] == cat].copy()
            if len(df_cat) < 5:
                continue

            # Crear features de tiempo
            df_cat['day_of_week'] = df_cat['date'].dt.dayofweek
            df_cat['month'] = df_cat['date'].dt.month
            df_cat['day_of_year'] = df_cat['date'].dt.dayofyear
            df_cat['is_weekend'] = (df_cat['day_of_week'] >= 5).astype(int)

            # Feature de lag (día anterior)
            df_cat['lag_1'] = df_cat['quantity'].shift(1).bfill()

            # Feature de media móvil 7 días
            df_cat['rolling_7'] = df_cat['quantity'].rolling(7, min_periods=1).mean().bfill()

            # Dividir en X e y
            features = ['day_of_week', 'month', 'day_of_year', 'is_weekend', 'lag_1', 'rolling_7']
            X = df_cat[features].values
            y = df_cat['quantity'].values

            # Entrenar modelo para esta categoría
            metrics = self.trainer.train(X, y, model_name=f"model_{cat}")

            # Guardar el mejor modelo global (el de menor MAPE)
            if 'mape' in metrics:
                if not best_overall_metrics or metrics['mape'] < best_overall_metrics.get('mape', float('inf')):
                    best_overall_metrics = metrics
                    self.trainer.save_best_model(BEST_MODEL_FILE)

        return best_overall_metrics