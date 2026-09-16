"""Referencia estacional transparente para series de demanda diaria."""

from __future__ import annotations

import numpy as np


class SeasonalAverageRegressor:
    """Pronostica con el promedio de las cuatro semanas equivalentes previas."""

    def __init__(self, feature_index: int):
        self.feature_index = feature_index

    def fit(self, _features: np.ndarray, _target: np.ndarray):
        return self

    def predict(self, features: np.ndarray) -> np.ndarray:
        values = np.asarray(features, dtype=float)[:, self.feature_index]
        return np.maximum(values, 0.0)
