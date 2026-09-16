"""Interfaz web para operar el pronóstico de demanda desde el navegador."""

from __future__ import annotations

import csv
import os
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

from application.services.catalog_service import CatalogService
from application.services.ingest_service import IngestService
from application.services.predict_service import PredictService
from application.services.train_service import TrainService
from config.settings import BASE_DIR, CATEGORY_PRODUCT_MIX_FILE, CATALOG_FILE, MODELS_FILE
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository
from infrastructure.repositories.csv_repository import CSVDemandRepository


ALLOWED_REPORT_EXTENSIONS = {".xlsx", ".xls", ".pdf"}
ALLOWED_CATALOG_EXTENSIONS = {".xlsx"}


def _parse_cutoff_date(value: Any) -> date | None:
    """Convierte la fecha base enviada por la interfaz a una fecha válida."""
    if value in (None, ""):
        return None
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").date()
    except (TypeError, ValueError) as exc:
        raise ValueError("La fecha base debe tener el formato AAAA-MM-DD.") from exc


def create_app(test_config: dict[str, Any] | None = None) -> Flask:
    """Crea la aplicación Flask sin ejecutar el servidor.

    Mantener esta fábrica permite reutilizar la interfaz en pruebas y evita que
    el servidor se inicie al importar este módulo.
    """
    app = Flask(__name__)
    app.config.from_mapping(
        MAX_CONTENT_LENGTH=20 * 1024 * 1024,
        CATALOG_FILE=CATALOG_FILE,
        SEND_FILE_MAX_AGE_DEFAULT=0,
    )
    if test_config:
        app.config.update(test_config)

    def catalog_service() -> CatalogService:
        return CatalogService(ExcelCatalogRepository(Path(app.config["CATALOG_FILE"])))

    @app.after_request
    def prevent_stale_dashboard(response):
        """Evita que el navegador conserve una interfaz JavaScript anterior."""
        if request.path == "/" or request.path == "/static/dashboard.js":
            response.headers["Cache-Control"] = "no-store, max-age=0"
        return response

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
            return _json_response(_build_dashboard(CSVDemandRepository(), catalog_service()))
        except Exception as exc:  # La respuesta debe ser útil para la interfaz.
            return _error_response(exc, 500)

    @app.get("/api/catalog")
    def active_catalog():
        try:
            service = catalog_service()
            products = service.get_all_products()
            return _json_response(
                {
                    "summary": service.get_summary(),
                    "products": [
                        {
                            "name": product.product_name,
                            "family": product.family,
                            "category": product.category,
                            "brand": product.brand,
                            "cost": round(product.cost, 2),
                        }
                        for product in products
                    ],
                }
            )
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except Exception as exc:
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
            active_catalog_service = catalog_service()
            ingest_service = IngestService(demand_repo, active_catalog_service)
            demands = ingest_service.process_file(str(temporary_path))

            return _json_response(
                {
                    "message": f"{filename} se procesó correctamente.",
                    "records": len(demands),
                    "report_dates": sorted({demand.date.strftime("%Y-%m-%d") for demand in demands}),
                    "save_summary": ingest_service.last_save_summary,
                    "catalog_summary": ingest_service.last_catalog_summary,
                    "dashboard": _build_dashboard(demand_repo, active_catalog_service),
                }
            )
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)

    @app.post("/api/catalog")
    def update_catalog():
        catalog = request.files.get("catalog")
        if catalog is None or not catalog.filename:
            return _error_response("Selecciona un catálogo Excel en formato .xlsx.", 400)

        filename = secure_filename(catalog.filename)
        suffix = Path(filename).suffix.lower()
        if suffix not in ALLOWED_CATALOG_EXTENSIONS:
            return _error_response("Formato no soportado. Usa un archivo .xlsx.", 400)

        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temporary_file:
                catalog.save(temporary_file)
                temporary_path = Path(temporary_file.name)

            # Se valida por completo antes de reemplazar el catálogo activo.
            ExcelCatalogRepository(temporary_path)
            target_path = Path(app.config["CATALOG_FILE"])
            target_path.parent.mkdir(parents=True, exist_ok=True)
            os.replace(temporary_path, target_path)
            temporary_path = None

            active_catalog_service = catalog_service()
            return _json_response(
                {
                    "message": f"{filename} se actualizó correctamente como catálogo maestro.",
                    "catalog": active_catalog_service.get_summary(),
                    "dashboard": _build_dashboard(CSVDemandRepository(), active_catalog_service),
                }
            )
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except PermissionError:
            return _error_response(
                "No se puede reemplazar el catálogo porque está abierto en otra aplicación. Ciérralo e inténtalo otra vez.",
                400,
            )
        except Exception as exc:
            return _error_response(exc, 500)
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)

    @app.post("/api/train")
    def train_model():
        try:
            body = request.get_json(silent=True) or {}
            cutoff_date = _parse_cutoff_date(body.get("cutoff_date"))
            service = TrainService(CSVDemandRepository())
            metrics = service.run(cutoff_date=cutoff_date)
            if not metrics:
                return _error_response("No hay suficientes datos por categoría para entrenar.", 400)
            mode_message = (
                f"Modelo de prueba histórica entrenado hasta el {cutoff_date.strftime('%d/%m/%Y')}."
                if cutoff_date
                else "Modelo entrenado con todo el historial disponible."
            )
            return _json_response(
                {
                    "message": (
                        f"{mode_message} Validación histórica para "
                        f"{service.last_training_summary['trained_categories']} categorías."
                    ),
                    "metrics": {name: round(float(value), 2) for name, value in metrics.items()},
                    "training_summary": service.last_training_summary,
                    "cutoff_date": cutoff_date.isoformat() if cutoff_date else None,
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
            cutoff_date = _parse_cutoff_date(body.get("cutoff_date"))
        except (TypeError, ValueError):
            return _error_response("Indica una cantidad válida de días y una fecha base válida.", 400)

        if not 1 <= days <= 90:
            return _error_response("Puedes pronosticar entre 1 y 90 días.", 400)
        if not MODELS_FILE.exists():
            return _error_response("Primero entrena el modelo para generar el pronóstico.", 400)

        try:
            service = PredictService(CSVDemandRepository())
            predictions = service.predict_future(days, cutoff_date=cutoff_date)
            products_by_category = _load_category_products()
            serialized_predictions = {
                category: [
                    {"date": demand.date.strftime("%Y-%m-%d"), "quantity": round(float(demand.quantity), 2)}
                    for demand in demands
                ]
                for category, demands in predictions.items()
            }
            return _json_response(
                {
                    "message": (
                        f"Pronóstico histórico generado desde el día posterior al {cutoff_date.strftime('%d/%m/%Y')} para {days} días."
                        if cutoff_date
                        else f"Pronóstico generado para los próximos {days} días."
                    ),
                    "days": days,
                    "cutoff_date": cutoff_date.isoformat() if cutoff_date else None,
                    "predictions": serialized_predictions,
                    "products_by_category": {
                        category: products_by_category.get(category, [])
                        for category in predictions
                    },
                    "product_forecasts_by_category": _build_product_forecasts(
                        serialized_predictions, products_by_category
                    ),
                    "validation_by_category": {
                        category: (service.predictor.metadata.get("validation_by_category") or {}).get(category, {})
                        for category in predictions
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


def _build_dashboard(demand_repo: CSVDemandRepository, catalog_service: CatalogService) -> dict[str, Any]:
    """Prepara datos simples y serializables para el tablero."""
    demands = demand_repo.get_all_demands()
    catalog_summary = catalog_service.get_summary()
    if not demands:
        return {
            "summary": {
                "records": 0,
                "categories": 0,
                "total_quantity": 0,
                "first_sale_date": None,
                "last_sale_date": None,
                "history_updated_at": None,
            },
            "categories": [],
            "recent": [],
            "catalog": {
                **catalog_summary,
                "historical_categories": 0,
                "mapped_historical_categories": 0,
                "unmapped_historical_categories": [],
            },
        }

    category_totals: dict[str, float] = {}
    for demand in demands:
        category_totals[demand.category] = category_totals.get(demand.category, 0) + float(demand.quantity)

    sorted_demands = sorted(demands, key=lambda demand: demand.date, reverse=True)
    active_categories = set(catalog_service.get_categories())
    total_quantity = sum(category_totals.values())
    categories = [
        {
            "name": name,
            "quantity": round(quantity, 2),
            "percentage": round((quantity / total_quantity) * 100, 2) if total_quantity else 0,
        }
        for name, quantity in sorted(category_totals.items(), key=lambda item: item[1], reverse=True)
    ]
    recent = [
        {
            "date": demand.date.strftime("%Y-%m-%d"),
            "category": demand.category,
            "quantity": round(float(demand.quantity), 2),
        }
        for demand in (demand for demand in sorted_demands if demand.quantity > 0)
    ][:12]
    return {
        "summary": {
            "records": len(demands),
            "categories": len(category_totals),
            "total_quantity": round(total_quantity, 2),
            "first_sale_date": sorted_demands[-1].date.strftime("%Y-%m-%d"),
            "last_sale_date": sorted_demands[0].date.strftime("%Y-%m-%d"),
            "history_updated_at": (
                demand_repo.get_last_updated_at().isoformat() if demand_repo.get_last_updated_at() else None
            ),
        },
        "categories": categories,
        "recent": recent,
        "catalog": {
            **catalog_summary,
            "historical_categories": len(category_totals),
            "mapped_historical_categories": len(set(category_totals) & active_categories),
            "unmapped_historical_categories": sorted(set(category_totals) - active_categories),
        },
    }


def _error_response(error: Exception | str, status: int):
    return _json_response({"error": str(error)}, status)


def _json_response(payload: dict[str, Any], status: int = 200):
    """Evita que el navegador muestre un resumen anterior desde su caché."""
    response = jsonify(payload)
    response.status_code = status
    response.headers["Cache-Control"] = "no-store"
    return response


def _load_category_products(max_products: int | None = None) -> dict[str, list[dict[str, float | int | str]]]:
    """Obtiene todos los productos y su participación histórica por categoría.

    ``max_products`` se conserva solo para consultas que quieran limitar el
    resultado. El pronóstico usa el valor predeterminado para no ocultar una
    parte de la categoría bajo una fila genérica.
    """
    if not CATEGORY_PRODUCT_MIX_FILE.exists():
        return {}
    product_totals: dict[str, dict[str, float]] = {}
    with CATEGORY_PRODUCT_MIX_FILE.open(encoding="utf-8", newline="") as source:
        for row in csv.DictReader(source):
            category = str(row.get("category", "")).strip()
            product = str(row.get("product", "")).strip()
            if not category or not product:
                continue
            try:
                quantity = float(row.get("historical_quantity", 0))
            except (TypeError, ValueError):
                continue
            category_products = product_totals.setdefault(category, {})
            category_products[product] = category_products.get(product, 0.0) + quantity

    rankings: dict[str, list[dict[str, float | int | str]]] = {}
    for category, products in product_totals.items():
        category_total = sum(products.values())
        if category_total <= 0:
            continue
        ordered_products = sorted(products.items(), key=lambda item: (-item[1], item[0]))
        if max_products is not None:
            ordered_products = ordered_products[:max_products]
        rankings[category] = [
            {
                "rank": rank,
                "name": product,
                "historical_quantity": round(quantity, 2),
                "historical_share": quantity / category_total,
            }
            for rank, (product, quantity) in enumerate(ordered_products, start=1)
        ]
    return rankings


def _build_product_forecasts(
    predictions: dict[str, list[dict[str, float | str]]],
    products_by_category: dict[str, list[dict[str, float | int | str]]],
) -> dict[str, list[dict[str, Any]]]:
    """Distribuye cada pronóstico de categoría según la mezcla histórica de ventas.

    Los productos no se entrenan como modelos independientes: esta asignación
    conserva el total validado de la categoría y muestra cómo se repartiría entre
    todos sus productos según su participación histórica.
    """
    result: dict[str, list[dict[str, Any]]] = {}
    for category, demands in predictions.items():
        ranking = products_by_category.get(category, [])
        total_share = sum(max(0.0, float(product["historical_share"])) for product in ranking)
        normalized_ranking = [
            (product, max(0.0, float(product["historical_share"])) / total_share)
            for product in ranking
        ] if total_share > 0 else []
        daily_forecasts: list[dict[str, Any]] = []
        for demand in demands:
            category_quantity = round(float(demand["quantity"]), 2)
            products: list[dict[str, Any]] = []
            allocated_quantity = 0.0
            for product, historical_share in normalized_ranking:
                quantity = round(category_quantity * historical_share, 2)
                products.append(
                    {
                        "name": product["name"],
                        "quantity": quantity,
                        "historical_share": round(historical_share * 100, 2),
                    }
                )
                allocated_quantity += quantity

            residual = round(category_quantity - allocated_quantity, 2)
            if residual and products:
                # La diferencia solo procede del redondeo a dos decimales. Se
                # asigna al producto con mayor participación para que el total
                # de productos coincida exactamente con el total de categoría.
                products[0]["quantity"] = round(products[0]["quantity"] + residual, 2)

            daily_forecasts.append(
                {
                    "date": demand["date"],
                    "category_quantity": category_quantity,
                    "products": products,
                }
            )
        result[category] = daily_forecasts
    return result
