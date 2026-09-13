import pandas as pd
import numpy as np
from datetime import timedelta
from typing import List, Dict

from domain.entities.demand import Demand
from domain.interfaces.repositories import DemandRepository
from application.services.history_fingerprint import history_fingerprint
from infrastructure.ml.model_predictor import ModelPredictor
from config.settings import MODELS_FILE, TIME_FEATURES

class PredictService:
    def __init__(self, demand_repo: DemandRepository):
        self.demand_repo = demand_repo
        self.predictor = ModelPredictor()

    def predict_future(self, days: int = 7) -> Dict[str, List[Demand]]:
        # 1. Obtener historial antes de cargar el modelo para comprobar que no
        #    haya quedado desactualizado después de una carga de reportes.
        demands = self.demand_repo.get_all_demands()
        if len({demand.date.date() for demand in demands}) < 30:
            raise ValueError("Se necesitan al menos 30 días de datos históricos")

        # 2. Cargar el modelo guardado
        try:
            self.predictor.load_models(MODELS_FILE)
        except FileNotFoundError:
            raise FileNotFoundError("Primero debes entrenar el modelo usando 'python main.py train'")
        if self.predictor.history_fingerprint != history_fingerprint(demands):
            raise ValueError("El historial cambió desde el último entrenamiento. Entrena el modelo nuevamente.")

        df = pd.DataFrame([
            {'date': d.date, 'category': d.category, 'quantity': d.quantity}
            for d in demands
        ])
        df = df.sort_values('date').reset_index(drop=True)

        categories = df['category'].unique()
        predictions_by_category = {}

        for cat in categories:
            # La selección es dinámica: solo se pronostican productos que el
            # entrenamiento anterior consideró aptos y para los que guardó un
            # modelo vigente. No se usa un mínimo fijo de filas aquí.
            if cat not in self.predictor.models:
                continue

            df_cat = df[df['category'] == cat].copy()
            if df_cat.empty:
                continue

            last_date = df_cat['date'].max()
            # Generar fechas futuras
            future_dates = [last_date + timedelta(days=i+1) for i in range(days)]

            # El modelo solo recibe variables disponibles antes de la fecha
            # futura. No se emplean ventas futuras ni medias que las incluyan.
            predictions = []
            for dt in future_dates:
                features = [
                    dt.weekday(),          # day_of_week
                    dt.month,              # month
                    dt.timetuple().tm_yday, # day_of_year
                    1 if dt.weekday() >= 5 else 0, # is_weekend
                ]
                prediction = float(self.predictor.predict(cat, np.array([features]))[0])
                prediction = max(0, round(prediction, 2))
                predictions.append(prediction)

            # Crear entidades Demand
            demands_future = [
                Demand(date=dt, category=cat, quantity=pred)
                for dt, pred in zip(future_dates, predictions)
            ]
            predictions_by_category[cat] = demands_future

        return predictions_by_category
