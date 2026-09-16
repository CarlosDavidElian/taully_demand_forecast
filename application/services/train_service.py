import pandas as pd
import numpy as np
import joblib
from datetime import datetime, timezone
from typing import Dict

from domain.interfaces.repositories import DemandRepository
from application.services.history_fingerprint import history_fingerprint
from infrastructure.ml.model_trainer import ModelTrainer
from config.settings import FORECAST_FEATURES, MAX_VALIDATION_WAPE, MODELS_FILE

class TrainService:
    MIN_ACTIVE_DAYS = 30

    def __init__(self, demand_repo: DemandRepository):
        self.demand_repo = demand_repo
        self.last_training_summary = {"trained_categories": 0, "excluded_categories": {}}

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

        # 3. Preparar datos por categoría. Los rezagos se desplazan antes de
        # entrenar, por lo que cada fila solo usa demanda conocida al cierre
        # del día anterior.
        categories = sorted(df['category'].unique())
        models_by_category = {}
        metrics_by_category = []
        excluded_categories = {}
        validation_by_category = {}

        for cat in categories:
            df_cat = df[df['category'] == cat].copy()
            active_days = int((df_cat['quantity'] > 0).sum())
            if active_days < self.MIN_ACTIVE_DAYS:
                excluded_categories[str(cat)] = (
                    f"menos de {self.MIN_ACTIVE_DAYS} días con demanda registrada ({active_days})"
                )
                continue

            training_frame = self._build_training_frame(df_cat)
            if len(training_frame) < 20:
                excluded_categories[str(cat)] = "no tiene suficientes observaciones después de crear rezagos"
                continue

            X = training_frame[FORECAST_FEATURES].values
            y = training_frame['quantity'].values

            trainer = ModelTrainer()
            metrics = trainer.train(
                X,
                y,
                seasonal_feature_index=FORECAST_FEATURES.index("seasonal_mean_4"),
                model_name=f"model_{cat}",
            )
            if not np.isfinite(metrics["wape"]) or metrics["wape"] > MAX_VALIDATION_WAPE:
                excluded_categories[str(cat)] = (
                    f"WAPE de validación {metrics['wape']:.1f}% supera el límite de {MAX_VALIDATION_WAPE:.0f}%"
                )
                continue
            models_by_category[str(cat)] = trainer.best_model
            metrics_by_category.append(metrics)
            validation_by_category[str(cat)] = {
                "method": trainer.best_model_type,
                "wape": round(float(metrics["wape"]), 2),
            }

        if not models_by_category:
            raise ValueError(
                "No hay productos entrenables: cada producto necesita al menos 30 fechas "
                "y cantidades que varíen. Revisa los reportes cargados."
            )

        # Guardar todos los modelos en un único artefacto simplifica la carga
        # desde la web y garantiza que cada pronóstico use su propia categoría.
        joblib.dump(
            {
                "version": 3,
                "models": models_by_category,
                "metadata": {
                    "history_fingerprint": history_fingerprint(demands),
                    "trained_at": datetime.now(timezone.utc).isoformat(),
                    "features": FORECAST_FEATURES,
                    "minimum_active_days": self.MIN_ACTIVE_DAYS,
                    "maximum_validation_wape": MAX_VALIDATION_WAPE,
                    "trained_categories": sorted(models_by_category),
                    "excluded_categories": excluded_categories,
                    "validation_by_category": validation_by_category,
                },
            },
            MODELS_FILE,
        )

        # La interfaz muestra un resumen representativo de las categorías
        # entrenadas, no las métricas del "mejor" modelo aislado.
        self.last_training_summary = {
            "trained_categories": len(models_by_category),
            "excluded_categories": excluded_categories,
            "validation_by_category": validation_by_category,
        }
        return {
            name: self._mean_valid_metric(metrics_by_category, name)
            for name in ("mae", "rmse", "wape")
        }

    @staticmethod
    def _build_training_frame(dataframe: pd.DataFrame) -> pd.DataFrame:
        frame = dataframe.sort_values('date').copy()
        frame['day_of_week'] = frame['date'].dt.dayofweek
        frame['month'] = frame['date'].dt.month
        frame['day_of_year'] = frame['date'].dt.dayofyear
        frame['is_weekend'] = (frame['day_of_week'] >= 5).astype(int)
        frame['lag_1'] = frame['quantity'].shift(1)
        frame['lag_7'] = frame['quantity'].shift(7)
        frame['rolling_mean_7'] = frame['quantity'].shift(1).rolling(window=7, min_periods=7).mean()
        frame['seasonal_mean_4'] = sum(frame['quantity'].shift(days) for days in (7, 14, 21, 28)) / 4
        return frame.dropna(subset=FORECAST_FEATURES)

    @staticmethod
    def _mean_valid_metric(metrics_by_category: list[dict[str, float]], metric: str) -> float:
        values = [item[metric] for item in metrics_by_category if np.isfinite(item[metric])]
        return float(np.mean(values)) if values else 0.0
