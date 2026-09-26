from pathlib import Path
import sys

import pandas as pd
from sklearn.model_selection import train_test_split

ROOT = Path(r"C:\taully_demand_forecast")
sys.path.insert(0, str(ROOT))

from application.services.train_service import TrainService
from config.settings import FORECAST_FEATURES
from infrastructure.ml.model_trainer import ModelTrainer


history = pd.read_csv(ROOT / "data" / "historial_demanda.csv")
history["date"] = pd.to_datetime(history["date"])

print("VALIDACIÓN HISTÓRICA DEL MODELO")
for category in sorted(history["category"].unique()):
    frame = TrainService._build_training_frame(history[history["category"] == category].copy())
    trainer = ModelTrainer()
    metrics = trainer.train(
        frame[FORECAST_FEATURES].values,
        frame["quantity"].values,
        seasonal_feature_index=FORECAST_FEATURES.index("seasonal_mean_4"),
        model_name=f"model_{category}",
    )
    _, X_test, _, y_test = train_test_split(
        frame[FORECAST_FEATURES].values,
        frame["quantity"].values,
        test_size=0.2,
        shuffle=False,
    )
    forecast_total = float(trainer.best_model.predict(X_test).sum())
    validation_dates = frame.iloc[-len(y_test):]["date"]
    print(
        f"{category}|{trainer.best_model_type}|{metrics['mape']:.2f}|{metrics['mae']:.2f}|{metrics['rmse']:.2f}|{metrics['wape']:.2f}|"
        f"{y_test.sum():.2f}|{forecast_total:.2f}|{len(y_test)}|{validation_dates.iloc[0].date()}|{validation_dates.iloc[-1].date()}"
    )
