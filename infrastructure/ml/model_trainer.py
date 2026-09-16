import joblib
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error
from sklearn.model_selection import train_test_split
from typing import Dict

from infrastructure.ml.seasonal_average_regressor import SeasonalAverageRegressor

class ModelTrainer:
    def __init__(self):
        self.best_model = None

    def train(
        self,
        X: np.ndarray,
        y: np.ndarray,
        seasonal_feature_index: int,
        model_name: str = "model",
    ) -> Dict[str, float]:
        # Dividir en entrenamiento y prueba (80/20)
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, shuffle=False)

        ml_model = RandomForestRegressor(n_estimators=300, random_state=42, min_samples_leaf=2)
        ml_model.fit(X_train, y_train)
        seasonal_model = SeasonalAverageRegressor(seasonal_feature_index).fit(X_train, y_train)

        candidates = {
            "random_forest": (ml_model, ml_model.predict(X_test)),
            "promedio_estacional": (seasonal_model, seasonal_model.predict(X_test)),
        }
        evaluated = {
            name: self._metrics(y_test, prediction)
            for name, (_model, prediction) in candidates.items()
        }
        self.best_model_type = min(evaluated, key=lambda name: evaluated[name]["wape"])
        self.best_model = candidates[self.best_model_type][0]
        return evaluated[self.best_model_type]

    @staticmethod
    def _metrics(y_test: np.ndarray, y_pred: np.ndarray) -> Dict[str, float]:

        mae = mean_absolute_error(y_test, y_pred)
        rmse = np.sqrt(mean_squared_error(y_test, y_pred))
        
        # MAPE (evitar divisiones por cero)
        mask = y_test != 0
        if np.any(mask):
            mape = np.mean(np.abs((y_test[mask] - y_pred[mask]) / y_test[mask])) * 100
        else:
            # En un tramo de prueba con demanda real igual a cero, el MAPE no
            # está definido. Se conserva como NaN para no convertirlo en una
            # precisión artificial al promediar categorías.
            mape = float('nan')

        denominator = np.abs(y_test).sum()
        wape = (np.abs(y_test - y_pred).sum() / denominator * 100) if denominator else float('inf')
        return {'mae': mae, 'rmse': rmse, 'mape': mape, 'wape': wape}

    def save_best_model(self, path):
        if self.best_model is not None:
            joblib.dump(self.best_model, path)
        else:
            raise ValueError("No hay modelo entrenado para guardar")
