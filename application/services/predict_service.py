import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import List, Dict

from domain.entities.demand import Demand
from domain.interfaces.repositories import DemandRepository
from infrastructure.ml.model_predictor import ModelPredictor
from config.settings import BEST_MODEL_FILE

class PredictService:
    def __init__(self, demand_repo: DemandRepository):
        self.demand_repo = demand_repo
        self.predictor = ModelPredictor()

    def predict_future(self, days: int = 7) -> Dict[str, List[Demand]]:
        # 1. Cargar el modelo guardado
        try:
            self.predictor.load_model(BEST_MODEL_FILE)
        except FileNotFoundError:
            raise FileNotFoundError("Primero debes entrenar el modelo usando 'python main.py train'")

        # 2. Obtener datos históricos para generar features futuros
        demands = self.demand_repo.get_all_demands()
        if len(demands) < 5:
            raise ValueError("Se necesitan al menos 30 días de datos históricos")

        df = pd.DataFrame([
            {'date': d.date, 'category': d.category, 'quantity': d.quantity}
            for d in demands
        ])
        df = df.sort_values('date').reset_index(drop=True)

        categories = df['category'].unique()
        predictions_by_category = {}

        for cat in categories:
            df_cat = df[df['category'] == cat].copy()
            if len(df_cat) < 5:
                continue

            last_date = df_cat['date'].max()
            last_quantity = df_cat['quantity'].iloc[-1]
            last_rolling = df_cat['quantity'].rolling(7, min_periods=1).mean().iloc[-1]

            # Generar fechas futuras
            future_dates = [last_date + timedelta(days=i+1) for i in range(days)]

            # Construir features para cada día futuro
            future_features = []
            for dt in future_dates:
                features = [
                    dt.weekday(),          # day_of_week
                    dt.month,              # month
                    dt.timetuple().tm_yday, # day_of_year
                    1 if dt.weekday() >= 5 else 0, # is_weekend
                    last_quantity,         # lag_1 (usamos la última conocida)
                    last_rolling           # rolling_7 (usamos la última conocida)
                ]
                future_features.append(features)

            X_future = np.array(future_features)
            predictions = self.predictor.predict(X_future)

            # Crear entidades Demand
            demands_future = [
                Demand(date=dt, category=cat, quantity=max(0, round(pred, 2)))  # No permitir negativos
                for dt, pred in zip(future_dates, predictions)
            ]
            predictions_by_category[cat] = demands_future

        return predictions_by_category