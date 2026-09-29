"""Comprueba por el mismo contrato web los resultados de corte histórico.

El archivo de modelo se redirige a una ruta temporal para no alterar el modelo
activo del usuario.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
import sys

sys.path.insert(0, str(ROOT))

from application.services import predict_service as predict_module
from application.services import train_service as train_module
from interfaces import web as web_module


CUTOFF = "2026-09-08"
EXPECTED = ROOT / ".tmp_anexo2" / "backtest_real_2026-09-09_a_2026-09-15.json"


def main():
    original = (train_module.MODELS_FILE, predict_module.MODELS_FILE, web_module.MODELS_FILE)
    with tempfile.TemporaryDirectory(prefix="taully-ui-verification-") as temp_dir:
        temporary_model = Path(temp_dir) / "models_by_category.pkl"
        train_module.MODELS_FILE = temporary_model
        predict_module.MODELS_FILE = temporary_model
        web_module.MODELS_FILE = temporary_model
        try:
            client = web_module.create_app({"TESTING": True}).test_client()
            trained = client.post("/api/train", json={"cutoff_date": CUTOFF})
            forecast = client.post("/api/forecast", json={"days": 7, "cutoff_date": CUTOFF})
        finally:
            train_module.MODELS_FILE, predict_module.MODELS_FILE, web_module.MODELS_FILE = original

    assert trained.status_code == 200, trained.get_json()
    assert forecast.status_code == 200, forecast.get_json()
    train_payload = trained.get_json()
    forecast_payload = forecast.get_json()
    reference = json.loads(EXPECTED.read_text(encoding="utf-8"))
    for category, detail in reference["categories"].items():
        reported = [item["quantity"] for item in forecast_payload["predictions"][category]]
        expected = detail["model_forecast"]
        assert reported == expected, f"{category}: {reported} != {expected}"

    result = {
        "cutoff": CUTOFF,
        "train_metrics_shown_in_program": train_payload["metrics"],
        "forecast_dates": [item["date"] for item in forecast_payload["predictions"]["ABARROTES"]],
        "categories": sorted(forecast_payload["predictions"]),
        "message": "Los pronósticos de la interfaz coinciden exactamente con la prueba histórica guardada.",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
