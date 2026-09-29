"""Prueba de reproducibilidad de la interfaz para 09--15/09/2026.

No modifica el historial, el catálogo ni el modelo persistente de la aplicación.
El modelo de prueba se guarda únicamente en un directorio temporal y se llama a
las mismas rutas Flask que utiliza la pantalla web.
"""

from __future__ import annotations

import json
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from application.services import predict_service as predict_module
from application.services import train_service as train_module
from application.services.catalog_service import CatalogService
from config import settings
from interfaces import web
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository


CUTOFF = date(2026, 9, 8)
HORIZON = 7
HISTORY_FILE = ROOT / "data" / "historial_demanda.csv"


def metric_values(actual: list[float], predicted: list[float]) -> dict[str, float]:
    observed = np.asarray(actual, dtype=float)
    forecast = np.asarray(predicted, dtype=float)
    nonzero = observed != 0
    return {
        "mae": float(mean_absolute_error(observed, forecast)),
        "rmse": float(np.sqrt(mean_squared_error(observed, forecast))),
        "mape": float(np.mean(np.abs((observed[nonzero] - forecast[nonzero]) / observed[nonzero])) * 100),
        "wape": float(np.abs(observed - forecast).sum() / np.abs(observed).sum() * 100),
    }


def pms7_recursive(history: list[float], horizon: int) -> list[float]:
    values = list(history)
    result: list[float] = []
    for _ in range(horizon):
        next_value = round(float(np.mean(values[-7:])), 2)
        result.append(next_value)
        values.append(next_value)
    return result


def rounded(metrics: dict[str, float]) -> dict[str, float]:
    return {key: round(value, 4) for key, value in metrics.items()}


def reconcile_source_reports(history: pd.DataFrame, expected_dates: list[date]) -> dict[str, object]:
    """Comprueba que las ventas fuente de los siete días coincidan con el CSV."""
    catalog = CatalogService(ExcelCatalogRepository(ROOT / "data" / "catalogo_maestro.xlsx"))
    reader = ExcelReader()
    days: list[dict[str, object]] = []
    all_match = True
    total_unmatched = 0.0
    for report_date in expected_dates:
        report = ROOT / "data" / f"reporte_ventas_{report_date.isoformat()}.xlsx"
        sales = reader.read_sales(str(report))
        from_report = {category: 0.0 for category in catalog.get_categories()}
        unmatched = 0.0
        for sale in sales:
            product = catalog.get_product(sale.product_name)
            if product is None:
                unmatched += float(sale.quantity)
            else:
                from_report[product.category] += float(sale.quantity)
        from_history = {
            category: float(
                history[(history["date"] == report_date) & (history["category"] == category)]["quantity"].iloc[0]
            )
            for category in sorted(history["category"].unique())
        }
        match = from_report == from_history
        all_match = all_match and match
        total_unmatched += unmatched
        days.append(
            {
                "date": report_date.isoformat(),
                "report": report.name,
                "source_category_totals": from_report,
                "history_category_totals": from_history,
                "unmatched_source_units": unmatched,
                "matches_history": match,
            }
        )
    return {"all_7_reports_match_history": all_match, "total_unmatched_source_units": total_unmatched, "days": days}


def reconcile_entire_history(history: pd.DataFrame) -> dict[str, object]:
    """Audita los reportes fuente completos sin escribir en el historial."""
    catalog = CatalogService(ExcelCatalogRepository(ROOT / "data" / "catalogo_maestro.xlsx"))
    categories = catalog.get_categories()
    reader = ExcelReader()
    reports = sorted((ROOT / "data").glob("reporte_ventas_*.xlsx"))
    errors: list[str] = []
    total_unmatched = 0.0
    seen_dates: set[date] = set()
    for report in reports:
        sales = reader.read_sales(str(report))
        dates_in_report = {sale.date.date() for sale in sales}
        if len(dates_in_report) != 1:
            errors.append(f"{report.name}: contiene {len(dates_in_report)} fechas")
            continue
        report_date = dates_in_report.pop()
        if report_date in seen_dates:
            errors.append(f"{report.name}: duplica la fecha {report_date.isoformat()}")
            continue
        seen_dates.add(report_date)
        report_totals = {category: 0.0 for category in categories}
        for sale in sales:
            product = catalog.get_product(sale.product_name)
            if product is None:
                total_unmatched += float(sale.quantity)
            else:
                report_totals[product.category] += float(sale.quantity)
        stored_rows = history[history["date"] == report_date]
        stored_totals = {
            category: float(stored_rows[stored_rows["category"] == category]["quantity"].iloc[0])
            if (stored_rows["category"] == category).any()
            else None
            for category in categories
        }
        if report_totals != stored_totals:
            errors.append(f"{report.name}: no coincide con historial")

    expected_dates = set(pd.date_range(history["date"].min(), history["date"].max(), freq="D").date)
    history_keys = history.groupby(["date", "category"]).size()
    duplicate_history_keys = int((history_keys > 1).sum())
    missing_history_dates = sorted(expected_dates - set(history["date"]))
    missing_report_dates = sorted(expected_dates - seen_dates)
    extra_report_dates = sorted(seen_dates - expected_dates)
    return {
        "history_rows": int(len(history)),
        "history_date_range": [history["date"].min().isoformat(), history["date"].max().isoformat()],
        "history_unique_dates": int(history["date"].nunique()),
        "history_categories": sorted(history["category"].unique().tolist()),
        "history_duplicate_date_category_keys": duplicate_history_keys,
        "source_report_files": len(reports),
        "source_report_unique_dates": len(seen_dates),
        "missing_history_dates": [day.isoformat() for day in missing_history_dates],
        "missing_report_dates": [day.isoformat() for day in missing_report_dates],
        "extra_report_dates": [day.isoformat() for day in extra_report_dates],
        "total_unmatched_source_units": total_unmatched,
        "mismatch_or_read_errors": errors,
        "all_source_reports_match_history": not errors and not missing_history_dates and not missing_report_dates and not extra_report_dates,
    }


def main() -> None:
    history = pd.read_csv(HISTORY_FILE)
    history["date"] = pd.to_datetime(history["date"]).dt.date
    expected_dates = [CUTOFF + timedelta(days=offset) for offset in range(1, HORIZON + 1)]
    end_date = expected_dates[-1]

    # Todas las rutas que escriben/leen el modelo apuntan al archivo temporal.
    original_train_model = train_module.MODELS_FILE
    original_predict_model = predict_module.MODELS_FILE
    original_web_model = web.MODELS_FILE
    with tempfile.TemporaryDirectory(prefix="taully-ui-repro-") as folder:
        temporary_model = Path(folder) / "models_by_category.pkl"
        train_module.MODELS_FILE = temporary_model
        predict_module.MODELS_FILE = temporary_model
        web.MODELS_FILE = temporary_model
        try:
            client = web.create_app({"TESTING": True}).test_client()
            train_response = client.post("/api/train", json={"cutoff_date": CUTOFF.isoformat()})
            forecast_response = client.post(
                "/api/forecast", json={"days": HORIZON, "cutoff_date": CUTOFF.isoformat()}
            )
        finally:
            train_module.MODELS_FILE = original_train_model
            predict_module.MODELS_FILE = original_predict_model
            web.MODELS_FILE = original_web_model

    if train_response.status_code != 200:
        raise RuntimeError(f"Entrenamiento UI falló: {train_response.status_code} {train_response.get_json()}")
    if forecast_response.status_code != 200:
        raise RuntimeError(f"Pronóstico UI falló: {forecast_response.status_code} {forecast_response.get_json()}")

    train_payload = train_response.get_json()
    forecast_payload = forecast_response.get_json()
    prediction_payload = forecast_payload["predictions"]

    categories: dict[str, dict] = {}
    for category in sorted(history["category"].unique()):
        before_cutoff = history[(history["category"] == category) & (history["date"] <= CUTOFF)].sort_values("date")
        observed_frame = history[
            (history["category"] == category)
            & (history["date"] > CUTOFF)
            & (history["date"] <= end_date)
        ].sort_values("date")
        actual_lookup = {row.date: float(row.quantity) for row in observed_frame.itertuples(index=False)}
        if set(actual_lookup) != set(expected_dates):
            raise RuntimeError(f"Fechas reales incompletas para {category}.")
        model_lookup = {
            date.fromisoformat(item["date"]): float(item["quantity"])
            for item in prediction_payload[category]
        }
        actual = [actual_lookup[day] for day in expected_dates]
        model = [model_lookup[day] for day in expected_dates]
        baseline = pms7_recursive(before_cutoff["quantity"].astype(float).tolist(), HORIZON)
        categories[category] = {
            "actual": actual,
            "pms7": baseline,
            "programa": model,
            "metricas_pms7": rounded(metric_values(actual, baseline)),
            "metricas_programa": rounded(metric_values(actual, model)),
            "modelo_seleccionado_por_el_programa": train_payload["training_summary"]["validation_by_category"][category],
        }

    actual_flat = [value for row in categories.values() for value in row["actual"]]
    pms_flat = [value for row in categories.values() for value in row["pms7"]]
    program_flat = [value for row in categories.values() for value in row["programa"]]
    total_actual = np.sum([row["actual"] for row in categories.values()], axis=0).tolist()
    total_pms = np.sum([row["pms7"] for row in categories.values()], axis=0).tolist()
    total_program = np.sum([row["programa"] for row in categories.values()], axis=0).tolist()

    result = {
        "ui_calls": {
            "training": {"endpoint": "/api/train", "cutoff_date": CUTOFF.isoformat(), "status": train_response.status_code},
            "forecast": {"endpoint": "/api/forecast", "cutoff_date": CUTOFF.isoformat(), "days": HORIZON, "status": forecast_response.status_code},
        },
        "period": {"training_from": str(history["date"].min()), "cutoff": CUTOFF.isoformat(), "forecast_from": expected_dates[0].isoformat(), "forecast_to": end_date.isoformat()},
        "ui_training_metrics": train_payload["metrics"],
        "categories": categories,
        "aggregate_35_category_day_observations": {
            "count": len(actual_flat),
            "metricas_pms7": rounded(metric_values(actual_flat, pms_flat)),
            "metricas_programa": rounded(metric_values(actual_flat, program_flat)),
        },
        "aggregate_daily_total_7_days": {
            "actual": total_actual,
            "pms7": total_pms,
            "programa": total_program,
            "metricas_pms7": rounded(metric_values(total_actual, total_pms)),
            "metricas_programa": rounded(metric_values(total_actual, total_program)),
        },
        "source_reconciliation": reconcile_source_reports(history, expected_dates),
    }
    if "--full-history" in sys.argv:
        result["full_history_reconciliation"] = reconcile_entire_history(history)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
