import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data"
MODELS_DIR = DATA_DIR / "models"

CATALOG_FILE = DATA_DIR / "catalogo_maestro.xlsx"
HISTORIAL_FILE = DATA_DIR / "historial_demanda.csv"
# Reporte de control con ventas que el catálogo aún no puede clasificar.
UNCATEGORIZED_SALES_FILE = DATA_DIR / "ventas_sin_categoria.csv"
CATEGORY_PRODUCT_MIX_FILE = DATA_DIR / "productos_por_categoria.csv"
HISTORY_BACKUPS_DIR = DATA_DIR / "backups"
# El usuario debe cargar un kardex actualizado desde la pantalla. El archivo
# de demostración se conserva solo para pruebas técnicas y nunca se selecciona
# automáticamente como evidencia de inventario.
INVENTORY_FILE = DATA_DIR / "inventario_producto.xlsx"
DEFAULT_INVENTORY_FILE = DATA_DIR / "inventario_producto_Taully_2026-04-01_a_2026-09-15.xlsx"
INVENTORY_UPLOADS_DIR = DATA_DIR / "inventarios_cargados"
INVENTORY_SELECTION_FILE = INVENTORY_UPLOADS_DIR / "inventario_activo.json"
POSTTEST_RESULTS_DIR = DATA_DIR / "posttests"
# Artefacto que contiene un modelo entrenado por cada categoría de producto.
MODELS_FILE = MODELS_DIR / "models_by_category.pkl"

TIME_FEATURES = ['day_of_week', 'month', 'day_of_year', 'is_weekend']
LAG_FEATURES = ['lag_1', 'lag_7', 'rolling_mean_7', 'seasonal_mean_4']
FORECAST_FEATURES = TIME_FEATURES + LAG_FEATURES
MAX_VALIDATION_WAPE = 50.0

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(MODELS_DIR, exist_ok=True)
