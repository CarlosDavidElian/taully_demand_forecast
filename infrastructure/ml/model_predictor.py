import joblib
import numpy as np
from typing import Any

class ModelPredictor:
    def __init__(self):
        self.models: dict[str, Any] = {}
        self.history_fingerprint: str | None = None

    def load_models(self, path: str):
        artifact = joblib.load(path)
        if not isinstance(artifact, dict) or not artifact.get("models"):
            raise ValueError("El modelo es antiguo o inválido. Entrénalo nuevamente.")
        metadata = artifact.get("metadata") or {}
        fingerprint = metadata.get("history_fingerprint")
        if not fingerprint:
            raise ValueError("El modelo no tiene trazabilidad del historial. Entrénalo nuevamente.")
        self.models = artifact["models"]
        self.history_fingerprint = str(fingerprint)

    def predict(self, category: str, X: np.ndarray) -> np.ndarray:
        model = self.models.get(category)
        if model is None:
            raise ValueError(f"No hay un modelo entrenado para la categoría '{category}'")
        return model.predict(X)
