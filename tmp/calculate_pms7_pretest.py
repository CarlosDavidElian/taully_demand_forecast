from pathlib import Path
import json

import numpy as np
import pandas as pd


history_path = Path(r"C:\taully_demand_forecast\data\historial_demanda.csv")
history = pd.read_csv(history_path, parse_dates=["date"])

cutoff = pd.Timestamp("2026-09-01")
period_start = pd.Timestamp("2026-09-02")
period_end = pd.Timestamp("2026-09-08")

results = {}
observations = []
for category in sorted(history["category"].unique()):
    training = history[(history["category"] == category) & (history["date"] <= cutoff)]
    actual = history[
        (history["category"] == category)
        & (history["date"] >= period_start)
        & (history["date"] <= period_end)
    ].sort_values("date")
    rolling = training.sort_values("date")["quantity"].astype(float).tail(7).tolist()
    predictions = []
    for _ in actual.itertuples():
        prediction = float(np.mean(rolling))
        predictions.append(prediction)
        rolling = rolling[1:] + [prediction]

    actual_values = actual["quantity"].to_numpy(dtype=float)
    prediction_values = np.array(predictions, dtype=float)
    error_values = np.abs(actual_values - prediction_values)
    result = {
        "actual_total": float(actual_values.sum()),
        "prediction_total": float(prediction_values.sum()),
        "mape": float(np.mean(error_values / actual_values) * 100),
        "mae": float(np.mean(error_values)),
        "rmse": float(np.sqrt(np.mean((actual_values - prediction_values) ** 2))),
        "wape": float(error_values.sum() / actual_values.sum() * 100),
    }
    results[category] = result
    observations.extend(
        {
            "date": row.date.strftime("%Y-%m-%d"),
            "category": category,
            "actual": float(row.quantity),
            "prediction": float(prediction),
        }
        for row, prediction in zip(actual.itertuples(), predictions)
    )

actual_all = np.array([row["actual"] for row in observations], dtype=float)
prediction_all = np.array([row["prediction"] for row in observations], dtype=float)
errors_all = np.abs(actual_all - prediction_all)
overall = {
    "actual_total": float(actual_all.sum()),
    "prediction_total": float(prediction_all.sum()),
    "mape": float(np.mean(errors_all / actual_all) * 100),
    "mae": float(np.mean(errors_all)),
    "rmse": float(np.sqrt(np.mean((actual_all - prediction_all) ** 2))),
    "wape": float(errors_all.sum() / actual_all.sum() * 100),
}

print(json.dumps({"cutoff": str(cutoff.date()), "period": [str(period_start.date()), str(period_end.date())], "by_category": results, "overall": overall}, ensure_ascii=False, indent=2))
