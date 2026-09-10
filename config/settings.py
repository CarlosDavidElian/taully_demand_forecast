import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data"
MODELS_DIR = DATA_DIR / "models"

CATALOG_FILE = DATA_DIR / "catalogo_maestro.xlsx"
HISTORIAL_FILE = DATA_DIR / "historial_demanda.csv"
# Artefacto que contiene un modelo entrenado por cada categoría de producto.
MODELS_FILE = MODELS_DIR / "models_by_category.pkl"

TIME_FEATURES = ['day_of_week', 'month', 'day_of_year', 'is_weekend']

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(MODELS_DIR, exist_ok=True)
