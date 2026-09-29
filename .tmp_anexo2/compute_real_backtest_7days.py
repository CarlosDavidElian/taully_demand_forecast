"""Calcula una comparación retrospectiva reproducible de siete días.

No modifica el modelo que utiliza la aplicación. Entrena el mismo servicio del
programa en un archivo temporal hasta la fecha de corte y compara sus
pronósticos con un promedio móvil simple de siete días sobre las mismas ventas
reales posteriores.
"""

from __future__ import annotations

import json
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error

ROOT = Path(__file__).resolve().parents[1]
import sys

sys.path.insert(0, str(ROOT))

from application.services import predict_service as predict_module
from application.services import train_service as train_module
from application.services.predict_service import PredictService
from application.services.train_service import TrainService
from infrastructure.repositories.csv_repository import CSVDemandRepository


CUTOFF = date(2026, 5, 18)
TEST_DAYS = 7
HISTORY = ROOT / "data" / "historial_demanda.csv"
OUTPUT = ROOT / ".tmp_anexo2" / "backtest_real_2026-05-19_a_2026-05-25.json"
TEMP_MODEL = ROOT / ".tmp_anexo2" / "models_backtest_2026-05-18.pkl"


def metrics(actual: list[float], predicted: list[float]) -> dict[str, float]:
    actual_values = np.array(actual, dtype=float)
    predicted_values = np.array(predicted, dtype=float)
    nonzero = actual_values != 0
    mape = float(np.mean(np.abs((actual_values[nonzero] - predicted_values[nonzero]) / actual_values[nonzero])) * 100) if nonzero.any() else None
    wape = float(np.abs(actual_values - predicted_values).sum() / np.abs(actual_values).sum() * 100) if np.abs(actual_values).sum() else None
    return {
        "mae": round(float(mean_absolute_error(actual_values, predicted_values)), 4),
        "rmse": round(float(np.sqrt(mean_squared_error(actual_values, predicted_values))), 4),
        "mape": round(mape, 4) if mape is not None else None,
        "wape": round(wape, 4) if wape is not None else None,
    }


def recursive_moving_average(values: list[float], days: int) -> list[float]:
    if len(values) < 7:
        raise ValueError("El promedio móvil simple de siete días requiere al menos siete ventas previas.")
    work = list(values)
    forecast: list[float] = []
    for _ in range(days):
        value = round(float(np.mean(work[-7:])), 2)
        forecast.append(value)
        work.append(value)
    return forecast


def main() -> None:
    end_date = CUTOFF + timedelta(days=TEST_DAYS)
    history = pd.read_csv(HISTORY)
    history["date"] = pd.to_datetime(history["date"]).dt.date
    expected_dates = [CUTOFF + timedelta(days=offset) for offset in range(1, TEST_DAYS + 1)]

    temporary_original_train = train_module.MODELS_FILE
    temporary_original_predict = predict_module.MODELS_FILE
    train_module.MODELS_FILE = TEMP_MODEL
    predict_module.MODELS_FILE = TEMP_MODEL
    try:
        repository = CSVDemandRepository(HISTORY)
        trainer = TrainService(repository)
        trainer.run(cutoff_date=CUTOFF)
        training_summary = trainer.last_training_summary
        model_predictions = PredictService(repository).predict_future(TEST_DAYS, cutoff_date=CUTOFF)
    finally:
        train_module.MODELS_FILE = temporary_original_train
        predict_module.MODELS_FILE = temporary_original_predict
        TEMP_MODEL.unlink(missing_ok=True)

    categories: dict[str, dict] = {}
    for category in sorted(history["category"].unique()):
        past = history[(history["category"] == category) & (history["date"] <= CUTOFF)].sort_values("date")
        actual_frame = history[(history["category"] == category) & (history["date"] > CUTOFF) & (history["date"] <= end_date)].sort_values("date")
        actual_by_date = {row.date: float(row.quantity) for row in actual_frame.itertuples(index=False)}
        missing = [value.isoformat() for value in expected_dates if value not in actual_by_date]
        if missing:
            raise ValueError(f"Faltan ventas reales para {category}: {', '.join(missing)}")
        baseline = recursive_moving_average([float(value) for value in past["quantity"]], TEST_DAYS)
        predicted = model_predictions.get(category)
        if predicted is None:
            raise ValueError(f"No se generó pronóstico del modelo para {category}.")
        model_by_date = {item.date.date(): float(item.quantity) for item in predicted}
        actual = [actual_by_date[value] for value in expected_dates]
        model = [model_by_date[value] for value in expected_dates]
        categories[category] = {
            "actual": actual,
            "baseline_pms_7": baseline,
            "model_forecast": model,
            "baseline_metrics": metrics(actual, baseline),
            "model_metrics": metrics(actual, model),
            "trained_model": training_summary["validation_by_category"].get(category, {}).get("method"),
            "training_wape": training_summary["validation_by_category"].get(category, {}).get("wape"),
        }

    total_actual = np.sum([item["actual"] for item in categories.values()], axis=0).tolist()
    total_baseline = np.sum([item["baseline_pms_7"] for item in categories.values()], axis=0).tolist()
    total_model = np.sum([item["model_forecast"] for item in categories.values()], axis=0).tolist()
    observation_actual = [value for item in categories.values() for value in item["actual"]]
    observation_baseline = [value for item in categories.values() for value in item["baseline_pms_7"]]
    observation_model = [value for item in categories.values() for value in item["model_forecast"]]
    result = {
        "method": {
            "pretest": "Promedio móvil simple recursivo de 7 días.",
            "postest": "Modelo seleccionado por el programa con entrenamiento hasta la fecha de corte.",
        },
        "cutoff_date": CUTOFF.isoformat(),
        "training_start": history["date"].min().isoformat(),
        "period_start": expected_dates[0].isoformat(),
        "period_end": expected_dates[-1].isoformat(),
        "dates": [value.isoformat() for value in expected_dates],
        "categories": categories,
        "overall": {
            "actual": total_actual,
            "baseline_pms_7": total_baseline,
            "model_forecast": total_model,
            "baseline_metrics": metrics(total_actual, total_baseline),
            "model_metrics": metrics(total_actual, total_model),
        },
        "all_category_day_observations": {
            "count": len(observation_actual),
            "baseline_metrics": metrics(observation_actual, observation_baseline),
            "model_metrics": metrics(observation_actual, observation_model),
        },
    }
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUTPUT)
    print(json.dumps(result["overall"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
