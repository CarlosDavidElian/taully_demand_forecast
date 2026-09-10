import pandas as pd
import numpy as np
from datetime import timedelta
from typing import List, Dict

from domain.entities.demand import Demand
from domain.interfaces.repositories import DemandRepository
from infrastructure.ml.model_predictor import ModelPredictor
from config.settings import MODELS_FILE

class PredictService:
    def __init__(self, demand_repo: DemandRepository):
        self.demand_repo = demand_repo
        self.predictor = ModelPredictor()

    def predict_future(self, days: int = 7) -> Dict[str, List[Demand]]:
        # 1. Cargar el modelo guardado
        try:
            self.predictor.load_models(MODELS_FILE)
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
            quantities = [float(quantity) for quantity in df_cat['quantity']]

            # Generar fechas futuras
            future_dates = [last_date + timedelta(days=i+1) for i in range(days)]

            # Predecir en secuencia: cada resultado alimenta el lag y la
            # media móvil del siguiente día.
            predictions = []
            for dt in future_dates:
                last_quantity = quantities[-1]
                last_rolling = float(np.mean(quantities[-7:]))
                features = [
                    dt.weekday(),          # day_of_week
                    dt.month,              # month
                    dt.timetuple().tm_yday, # day_of_year
                    1 if dt.weekday() >= 5 else 0, # is_weekend
                    last_quantity,         # lag_1 (usamos la última conocida)
                    last_rolling           # rolling_7 (usamos la última conocida)
                ]
                prediction = float(self.predictor.predict(cat, np.array([features]))[0])
                prediction = max(0, round(prediction, 2))
                predictions.append(prediction)
                quantities.append(prediction)

            # Crear entidades Demand
            demands_future = [
                Demand(date=dt, category=cat, quantity=pred)
                for dt, pred in zip(future_dates, predictions)
            ]
            predictions_by_category[cat] = demands_future

        return predictions_by_category
