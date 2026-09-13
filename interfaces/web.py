"""Interfaz web para operar el pronóstico de demanda desde el navegador."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

from application.services.ingest_service import IngestService
from application.services.predict_service import PredictService
from application.services.train_service import TrainService
from config.settings import BASE_DIR, MODELS_FILE
from infrastructure.repositories.csv_repository import CSVDemandRepository


ALLOWED_REPORT_EXTENSIONS = {".xlsx", ".xls", ".pdf"}


def create_app(test_config: dict[str, Any] | None = None) -> Flask:
    """Crea la aplicación Flask sin ejecutar el servidor.

    Mantener esta fábrica permite reutilizar la interfaz en pruebas y evita que
    el servidor se inicie al importar este módulo.
    """
    app = Flask(__name__)
    app.config.from_mapping(MAX_CONTENT_LENGTH=20 * 1024 * 1024)
    if test_config:
        app.config.update(test_config)

    @app.get("/")
    def dashboard() -> str:
        return render_template("dashboard.html")

    @app.get("/logo.jpg")
    def brand_logo():
        """Entrega el logo corporativo que se muestra en el encabezado."""
        return send_from_directory(str(BASE_DIR), "logo.jpg")

    @app.get("/favicon.ico")
    def favicon():
        """Entrega el formato estándar que los navegadores buscan en una pestaña."""
        static_dir = BASE_DIR / "interfaces" / "static"
        return send_from_directory(str(static_dir), "taully-favicon.ico", mimetype="image/vnd.microsoft.icon")

    @app.get("/api/dashboard")
    def dashboard_data():
        try:
            return _json_response(_build_dashboard(CSVDemandRepository()))
        except Exception as exc:  # La respuesta debe ser útil para la interfaz.
            return _error_response(exc, 500)

    @app.post("/api/reports")
    def load_report():
        report = request.files.get("report")
        if report is None or not report.filename:
            return _error_response("Selecciona un reporte Excel o PDF.", 400)

        filename = secure_filename(report.filename)
        suffix = Path(filename).suffix.lower()
        if suffix not in ALLOWED_REPORT_EXTENSIONS:
            return _error_response("Formato no soportado. Usa .xlsx, .xls o .pdf.", 400)

        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temporary_file:
                report.save(temporary_file)
                temporary_path = Path(temporary_file.name)

            demand_repo = CSVDemandRepository()
            ingest_service = IngestService(demand_repo)
            demands = ingest_service.process_file(str(temporary_path))

            return _json_response(
                {
                    "message": f"{filename} se procesó correctamente.",
                    "records": len(demands),
                    "save_summary": ingest_service.last_save_summary,
                    "dashboard": _build_dashboard(demand_repo),
                }
            )
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)

    @app.post("/api/train")
    def train_model():
        try:
            metrics = TrainService(CSVDemandRepository()).run()
            if not metrics:
                return _error_response("No hay suficientes datos por categoría para entrenar.", 400)
            return _json_response(
                {
                    "message": "Modelo entrenado correctamente.",
                    "metrics": {name: round(float(value), 2) for name, value in metrics.items()},
                }
            )
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)

    @app.post("/api/forecast")
    def forecast():
        body = request.get_json(silent=True) or {}
        try:
            days = int(body.get("days", 7))
        except (TypeError, ValueError):
            return _error_response("Indica una cantidad válida de días.", 400)

        if not 1 <= days <= 90:
            return _error_response("Puedes pronosticar entre 1 y 90 días.", 400)
        if not MODELS_FILE.exists():
            return _error_response("Primero entrena el modelo para generar el pronóstico.", 400)

        try:
            predictions = PredictService(CSVDemandRepository()).predict_future(days)
            return _json_response(
                {
                    "message": f"Pronóstico generado para los próximos {days} días.",
                    "days": days,
                    "predictions": {
                        category: [
                            {"date": demand.date.strftime("%Y-%m-%d"), "quantity": round(float(demand.quantity), 2)}
                            for demand in demands
                        ]
                        for category, demands in predictions.items()
                    },
                }
            )
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)

    @app.errorhandler(413)
    def file_too_large(_error):
        return _error_response("El reporte supera el límite de 20 MB.", 413)

    return app


def _build_dashboard(demand_repo: CSVDemandRepository) -> dict[str, Any]:
    """Prepara datos simples y serializables para el tablero."""
    demands = demand_repo.get_all_demands()
    if not demands:
        return {
            "summary": {
                "records": 0,
                "products": 0,
                "total_quantity": 0,
                "last_sale_date": None,
                "history_updated_at": None,
            },
            "products": [],
            "recent": [],
        }

    product_totals: dict[str, float] = {}
    for demand in demands:
        product_totals[demand.category] = product_totals.get(demand.category, 0) + float(demand.quantity)

    sorted_demands = sorted(demands, key=lambda demand: demand.date, reverse=True)
    products = [
        {"name": name, "quantity": round(quantity, 2)}
        for name, quantity in sorted(product_totals.items(), key=lambda item: item[1], reverse=True)
    ]
    recent = [
        {
            "date": demand.date.strftime("%Y-%m-%d"),
            "category": demand.category,
            "quantity": round(float(demand.quantity), 2),
        }
        for demand in sorted_demands[:12]
    ]
    return {
        "summary": {
            "records": len(demands),
            "products": len(product_totals),
            "total_quantity": round(sum(product_totals.values()), 2),
            "last_sale_date": sorted_demands[0].date.strftime("%Y-%m-%d"),
            "history_updated_at": (
                demand_repo.get_last_updated_at().isoformat() if demand_repo.get_last_updated_at() else None
            ),
        },
        "products": products,
        "recent": recent,
    }


def _error_response(error: Exception | str, status: int):
    return _json_response({"error": str(error)}, status)


def _json_response(payload: dict[str, Any], status: int = 200):
    """Evita que el navegador muestre un resumen anterior desde su caché."""
    response = jsonify(payload)
    response.status_code = status
    response.headers["Cache-Control"] = "no-store"
    return response
