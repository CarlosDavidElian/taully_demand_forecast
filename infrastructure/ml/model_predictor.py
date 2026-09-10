import joblib
import numpy as np
from typing import Any

class ModelPredictor:
    def __init__(self):
        self.models: dict[str, Any] = {}

    def load_models(self, path: str):
        models = joblib.load(path)
        if not isinstance(models, dict) or not models:
            raise ValueError("El archivo de modelos no contiene categorías entrenadas")
        self.models = models

    def predict(self, category: str, X: np.ndarray) -> np.ndarray:
        model = self.models.get(category)
        if model is None:
            raise ValueError(f"No hay un modelo entrenado para la categoría '{category}'")
        return model.predict(X)
