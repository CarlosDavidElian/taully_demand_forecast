"""Backtesting semanal reproducible con todo el historial disponible.

Evalúa dos métodos sobre las mismas semanas reales, desde el 20 de mayo al
15 de septiembre de 2026. El modelo se entrena de nuevo antes de cada semana
y usa los mismos servicios de entrenamiento y predicción de la aplicación.
El artefacto del modelo se guarda solo en una carpeta temporal.
"""

from __future__ import annotations

import json
import tempfile
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

HISTORY = ROOT / "data" / "historial_demanda.csv"
OUTPUT = ROOT / ".tmp_anexo2" / "validacion_rodante_2026-05-20_a_2026-09-15.json"
FIRST_CUTOFF = date(2026, 5, 19)
LAST_CUTOFF = date(2026, 9, 8)
HORIZON = 7


def calculate_metrics(actual: list[float], predicted: list[float]) -> dict[str, float]:
    actual_array = np.asarray(actual, dtype=float)
    predicted_array = np.asarray(predicted, dtype=float)
    nonzero = actual_array != 0
    return {
        "mae": round(float(mean_absolute_error(actual_array, predicted_array)), 6),
        "rmse": round(float(np.sqrt(mean_squared_error(actual_array, predicted_array))), 6),
        "mape": round(float(np.mean(np.abs((actual_array[nonzero] - predicted_array[nonzero]) / actual_array[nonzero])) * 100), 6),
        "wape": round(float(np.abs(actual_array - predicted_array).sum() / np.abs(actual_array).sum() * 100), 6),
    }


def recursive_pms_7(values: list[float]) -> list[float]:
    current = list(values)
    result = []
    for _ in range(HORIZON):
        next_value = round(float(np.mean(current[-7:])), 2)
        result.append(next_value)
        current.append(next_value)
    return result


def cutoff_dates() -> list[date]:
    result = []
    current = FIRST_CUTOFF
    while current <= LAST_CUTOFF:
        result.append(current)
        current += timedelta(days=HORIZON)
    assert result[-1] == LAST_CUTOFF
    return result


def main() -> None:
    history = pd.read_csv(HISTORY)
    history["date"] = pd.to_datetime(history["date"]).dt.date
    source_start = history["date"].min()
    source_end = history["date"].max()
    categories = sorted(history["category"].unique())
    all_actual: list[float] = []
    all_pms: list[float] = []
    all_model: list[float] = []
    by_category = {category: {"actual": [], "pms": [], "model": []} for category in categories}
    folds: list[dict] = []

    original_train_model = train_module.MODELS_FILE
    original_predict_model = predict_module.MODELS_FILE
    try:
        with tempfile.TemporaryDirectory(prefix="taully-rolling-backtest-") as temporary_directory:
            repository = CSVDemandRepository(HISTORY)
            for index, cutoff in enumerate(cutoff_dates(), start=1):
                temp_model = Path(temporary_directory) / f"models_{cutoff.isoformat()}.pkl"
                train_module.MODELS_FILE = temp_model
                predict_module.MODELS_FILE = temp_model
                trainer = TrainService(repository)
                trainer.run(cutoff_date=cutoff)
                prediction = PredictService(repository).predict_future(HORIZON, cutoff_date=cutoff)
                test_dates = [cutoff + timedelta(days=step) for step in range(1, HORIZON + 1)]
                fold_rows: list[dict] = []
                fold_actual: list[float] = []
                fold_pms: list[float] = []
                fold_model: list[float] = []

                for category in categories:
                    past = history[(history["category"] == category) & (history["date"] <= cutoff)].sort_values("date")
                    observed = history[(history["category"] == category) & (history["date"] > cutoff) & (history["date"] <= test_dates[-1])].sort_values("date")
                    observed_by_date = {row.date: float(row.quantity) for row in observed.itertuples(index=False)}
                    assert all(day in observed_by_date for day in test_dates), f"faltan observaciones para {category} tras {cutoff}"
                    actual = [observed_by_date[day] for day in test_dates]
                    pms = recursive_pms_7([float(value) for value in past["quantity"]])
                    model = [float(item.quantity) for item in prediction[category]]
                    assert len(model) == HORIZON
                    all_actual.extend(actual)
                    all_pms.extend(pms)
                    all_model.extend(model)
                    fold_actual.extend(actual)
                    fold_pms.extend(pms)
                    fold_model.extend(model)
                    by_category[category]["actual"].extend(actual)
                    by_category[category]["pms"].extend(pms)
                    by_category[category]["model"].extend(model)
                    fold_rows.append(
                        {
                            "category": category,
                            "actual_total": round(sum(actual), 2),
                            "pms_total": round(sum(pms), 2),
                            "model_total": round(sum(model), 2),
                            "pms_metrics": calculate_metrics(actual, pms),
                            "model_metrics": calculate_metrics(actual, model),
                            "selected_method": trainer.last_training_summary["validation_by_category"][category]["method"],
                            "internal_validation_wape": trainer.last_training_summary["validation_by_category"][category]["wape"],
                        }
                    )

                folds.append(
                    {
                        "number": index,
                        "cutoff_date": cutoff.isoformat(),
                        "test_start": test_dates[0].isoformat(),
                        "test_end": test_dates[-1].isoformat(),
                        "global_metrics": {
                            "pms": calculate_metrics(fold_actual, fold_pms),
                            "model": calculate_metrics(fold_actual, fold_model),
                        },
                        "categories": fold_rows,
                    }
                )
                temp_model.unlink(missing_ok=True)
    finally:
        train_module.MODELS_FILE = original_train_model
        predict_module.MODELS_FILE = original_predict_model

    result = {
        "design": {
            "source_period": {"start": source_start.isoformat(), "end": source_end.isoformat()},
            "first_cutoff": FIRST_CUTOFF.isoformat(),
            "last_cutoff": LAST_CUTOFF.isoformat(),
            "test_period": {"start": (FIRST_CUTOFF + timedelta(days=1)).isoformat(), "end": (LAST_CUTOFF + timedelta(days=HORIZON)).isoformat()},
            "folds": len(folds),
            "horizon_days": HORIZON,
            "observations": len(all_actual),
            "pretest_method": "Promedio móvil simple recursivo de siete días (PMS-7).",
            "postest_method": "Modelo elegido por la aplicación tras entrenarse hasta cada fecha de corte.",
        },
        "overall": {
            "pms": calculate_metrics(all_actual, all_pms),
            "model": calculate_metrics(all_actual, all_model),
            "actual_total": round(sum(all_actual), 2),
            "pms_total": round(sum(all_pms), 2),
            "model_total": round(sum(all_model), 2),
        },
        "by_category": {
            category: {
                "observations": len(values["actual"]),
                "actual_total": round(sum(values["actual"]), 2),
                "pms_total": round(sum(values["pms"]), 2),
                "model_total": round(sum(values["model"]), 2),
                "pms": calculate_metrics(values["actual"], values["pms"]),
                "model": calculate_metrics(values["actual"], values["model"]),
            }
            for category, values in by_category.items()
        },
        "folds": folds,
    }
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUTPUT)
    print(json.dumps(result["overall"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
