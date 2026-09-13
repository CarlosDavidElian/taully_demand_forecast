import pandas as pd
import numpy as np
import joblib
from datetime import datetime, timezone
from typing import Dict

from domain.interfaces.repositories import DemandRepository
from application.services.history_fingerprint import history_fingerprint
from infrastructure.ml.model_trainer import ModelTrainer
from config.settings import MODELS_FILE, TIME_FEATURES

class TrainService:
    def __init__(self, demand_repo: DemandRepository):
        self.demand_repo = demand_repo

    def run(self) -> Dict[str, float]:
        # 1. Obtener historial
        demands = self.demand_repo.get_all_demands()
        if len({demand.date.date() for demand in demands}) < 30:
            raise ValueError("Se necesitan al menos 30 días de datos históricos para entrenar")

        # 2. Convertir a DataFrame y crear features temporales
        df = pd.DataFrame([
            {'date': d.date, 'category': d.category, 'quantity': d.quantity}
            for d in demands
        ])
        df = df.sort_values('date').reset_index(drop=True)

        # 3. Preparar datos por producto. Se usan solamente variables de
        # calendario: las antiguas variables lag/media móvil incluían la venta
        # que se quería predecir y producían métricas engañosamente perfectas.
        categories = df['category'].unique()
        models_by_category = {}
        metrics_by_category = []

        for cat in categories:
            df_cat = df[df['category'] == cat].copy()
            if df_cat['date'].nunique() < 30:
                continue
            if df_cat['quantity'].nunique() < 2:
                # Una serie constante no permite demostrar precisión de un
                # pronóstico; se informa al usuario en vez de guardar un
                # modelo con MAPE artificialmente igual a cero.
                continue

            # Crear features de tiempo
            df_cat['day_of_week'] = df_cat['date'].dt.dayofweek
            df_cat['month'] = df_cat['date'].dt.month
            df_cat['day_of_year'] = df_cat['date'].dt.dayofyear
            df_cat['is_weekend'] = (df_cat['day_of_week'] >= 5).astype(int)

            # Dividir en X e y
            features = TIME_FEATURES
            X = df_cat[features].values
            y = df_cat['quantity'].values

            # Cada categoría necesita su propio modelo: sus cantidades y su
            # comportamiento histórico no son intercambiables.
            trainer = ModelTrainer()
            metrics = trainer.train(X, y, model_name=f"model_{cat}")
            models_by_category[str(cat)] = trainer.best_model
            metrics_by_category.append(metrics)

        if not models_by_category:
            raise ValueError(
                "No hay productos entrenables: cada producto necesita al menos 30 fechas "
                "y cantidades que varíen. Revisa los reportes cargados."
            )

        # Guardar todos los modelos en un único artefacto simplifica la carga
        # desde la web y garantiza que cada pronóstico use su propia categoría.
        joblib.dump(
            {
                "version": 2,
                "models": models_by_category,
                "metadata": {
                    "history_fingerprint": history_fingerprint(demands),
                    "trained_at": datetime.now(timezone.utc).isoformat(),
                    "features": features,
                },
            },
            MODELS_FILE,
        )

        # La interfaz muestra un resumen representativo de las categorías
        # entrenadas, no las métricas del "mejor" modelo aislado.
        return {
            name: float(np.mean([metrics[name] for metrics in metrics_by_category]))
            for name in ("mae", "rmse", "mape")
        }
