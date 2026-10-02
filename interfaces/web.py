"""Interfaz web para operar el pronóstico de demanda desde el navegador."""

from __future__ import annotations

import csv
import hashlib
import json
import math
import os
import tempfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template, request, send_file, send_from_directory
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from werkzeug.utils import secure_filename

from application.services.catalog_service import CatalogService
from application.services.ingest_service import IngestService
from application.services.inventory_service import InventoryService
from application.services.predict_service import PredictService
from application.services.posttest_evaluation_service import PosttestEvaluationService
from application.services.train_service import TrainService
from config.settings import (
    BASE_DIR,
    CATEGORY_PRODUCT_MIX_FILE,
    CATALOG_FILE,
    DATA_DIR,
    DEFAULT_INVENTORY_FILE,
    INVENTORY_FILE,
    INVENTORY_SELECTION_FILE,
    INVENTORY_UPLOADS_DIR,
    MODELS_FILE,
    POSTTEST_RESULTS_DIR,
)
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.repositories.csv_repository import CSVDemandRepository
from infrastructure.repositories.posttest_result_repository import PosttestResultRepository


ALLOWED_REPORT_EXTENSIONS = {".xlsx", ".xls", ".pdf"}
ALLOWED_CATALOG_EXTENSIONS = {".xlsx"}
ALLOWED_INVENTORY_EXTENSIONS = {".xlsx"}


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
        INVENTORY_FILE=INVENTORY_FILE,
        DEFAULT_INVENTORY_FILE=DEFAULT_INVENTORY_FILE,
        INVENTORY_UPLOADS_DIR=INVENTORY_UPLOADS_DIR,
        INVENTORY_SELECTION_FILE=INVENTORY_SELECTION_FILE,
        POSTTEST_RESULTS_DIR=POSTTEST_RESULTS_DIR,
        SEND_FILE_MAX_AGE_DEFAULT=0,
    )
    if test_config:
        app.config.update(test_config)

    def catalog_service() -> CatalogService:
        return CatalogService(ExcelCatalogRepository(Path(app.config["CATALOG_FILE"])))

    def _file_sha256(path: Path) -> str | None:
        """Obtiene una huella de archivo para trazabilidad sin exponer rutas."""
        if not path.is_file():
            return None
        try:
            digest = hashlib.sha256()
            with path.open("rb") as source:
                for block in iter(lambda: source.read(1024 * 1024), b""):
                    digest.update(block)
            return digest.hexdigest()
        except OSError:
            return None

    def _inventory_source_metadata(
        path: Path,
        source_id: str | None,
        *,
        configured: bool = False,
    ) -> dict[str, str | bool | None]:
        """Describe de manera honesta el origen del archivo de inventario."""
        demonstration_id = _file_sha256(Path(app.config["DEFAULT_INVENTORY_FILE"]))
        is_demonstration = bool(
            demonstration_id
            and source_id
            and source_id == demonstration_id
        )
        if is_demonstration:
            return {
                "source_kind": "archivo de demostración del proyecto",
                "evidence_status": "no_apto_para_evidencia",
                "is_demonstration": True,
            }
        return {
            "source_kind": "archivo de inventario configurado" if configured else "archivo cargado por el usuario",
            "evidence_status": "requiere_verificacion_de_fuente",
            "is_demonstration": False,
        }

    def active_inventory() -> dict[str, Any] | None:
        """Obtiene el inventario activo sin seleccionar la demostración por defecto."""
        uploads_directory = Path(app.config["INVENTORY_UPLOADS_DIR"])
        selection_file = Path(app.config["INVENTORY_SELECTION_FILE"])
        default_selection_file = Path(INVENTORY_SELECTION_FILE)

        # En pruebas automatizadas, ignoramos la selección persistida del proyecto
        # para que el resultado no dependa de un archivo cargado previamente en
        # el entorno local del desarrollador.
        if app.config.get("TESTING") and selection_file.resolve() == default_selection_file.resolve():
            selection_file = None

        if selection_file is not None and selection_file.is_file():
            try:
                selection = json.loads(selection_file.read_text(encoding="utf-8"))
                stored_name = Path(str(selection.get("stored_name", ""))).name
                candidate = uploads_directory / stored_name
                if stored_name and candidate.is_file():
                    source_id = str(selection.get("source_id") or candidate.stem)
                    return {
                        "path": candidate,
                        "name": str(selection.get("original_name") or candidate.name),
                        "source_id": source_id,
                        "uploaded_at": selection.get("updated_at"),
                        **_inventory_source_metadata(candidate, source_id),
                    }
            except (OSError, TypeError, ValueError, json.JSONDecodeError):
                # Una referencia antigua o incompleta no impide usar el archivo
                # cargado ni volver a cargar un inventario.
                pass

        # Solo se permite una fuente configurada explícitamente. El Excel de
        # demostración incluido con el proyecto no se usa de forma automática.
        candidate = Path(app.config["INVENTORY_FILE"])
        if candidate.is_file():
            source_id = _file_sha256(candidate)
            return {
                "path": candidate,
                "name": candidate.name,
                "source_id": source_id,
                "uploaded_at": None,
                **_inventory_source_metadata(candidate, source_id, configured=True),
            }
        return None

    def set_active_inventory(stored_path: Path, original_name: str, source_id: str | None) -> None:
        """Conserva una referencia pequeña; el Excel original queda inmutable."""
        uploads_directory = Path(app.config["INVENTORY_UPLOADS_DIR"])
        selection_file = Path(app.config["INVENTORY_SELECTION_FILE"])
        uploads_directory.mkdir(parents=True, exist_ok=True)
        if stored_path.parent.resolve() != uploads_directory.resolve():
            raise ValueError("La copia del inventario no se guardó en la ubicación esperada.")
        payload = {
            "stored_name": stored_path.name,
            "original_name": original_name,
            "source_id": source_id or stored_path.stem,
            "updated_at": datetime.now().isoformat(),
        }
        temporary_path = selection_file.with_suffix(".tmp")
        temporary_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        os.replace(temporary_path, selection_file)

    def inventory_payload(
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> dict[str, Any]:
        """Resume los indicadores de inventario del archivo activo."""
        source = active_inventory()
        if source is None:
            return {
                "available": False,
                "message": "Carga el kardex real de la tienda para calcular ISI y TQS.",
                "metrics": [],
            }

        source_path = Path(source["path"])
        service = InventoryService()
        imported = service.import_inventory(source_path)
        metrics = service.calculate_category_metrics(source_path, start_date, end_date)
        reconciliation = service.reconciliation_summary(source_path, start_date, end_date)
        incomplete_coverage = [
            metric.category for metric in metrics if metric.coverage_rate < 100 - 1e-6
        ]
        warnings: list[str] = []
        if source["is_demonstration"]:
            warnings.append(
                "Este archivo es una demostración técnica del proyecto; no sustituye el kardex real ni debe usarse como evidencia de tesis."
            )
        if not reconciliation["reconciled"]:
            warnings.append(
                "Las ventas registradas y el balance de stock no cierran completamente; revisa ajustes, devoluciones, mermas y reposiciones antes de usar el indicador como evidencia."
            )
        if incomplete_coverage:
            warnings.append(
                "Hay cobertura incompleta de fechas en: " + ", ".join(incomplete_coverage) + "."
            )
        source_kind = str(source["source_kind"])
        message = (
            "Se muestran indicadores del archivo de demostración. Carga el kardex real de la tienda para resultados oficiales."
            if source["is_demonstration"]
            else "Indicadores calculados con el archivo de inventario activo. Verifica que sea el kardex oficial de la tienda."
        )
        return {
            "available": True,
            "message": message,
            "inventory": {
                "name": str(source["name"] or imported.original_name),
                "source_id": source["source_id"],
                "source_kind": source_kind,
                "evidence_status": source["evidence_status"],
                "is_demonstration": bool(source["is_demonstration"]),
                "uploaded_at": source["uploaded_at"],
                "records": imported.records,
                "categories": imported.categories,
                "start_date": imported.start_date.isoformat(),
                "end_date": imported.end_date.isoformat(),
            },
            "metrics": [metric.to_dict() for metric in metrics],
            "reconciliation": reconciliation,
            "warnings": warnings,
        }

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

    @app.get("/api/inventory")
    def inventory_data():
        """Muestra los indicadores calculados para el inventario activo."""
        try:
            return _json_response(inventory_payload())
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)

    @app.post("/api/inventory")
    def load_inventory():
        """Valida y conserva un kardex Excel sin reemplazar cargas anteriores."""
        inventory = request.files.get("inventory")
        if inventory is None or not inventory.filename:
            return _error_response("Selecciona un archivo de inventario en formato .xlsx.", 400)

        filename = secure_filename(inventory.filename)
        suffix = Path(filename).suffix.lower()
        if suffix not in ALLOWED_INVENTORY_EXTENSIONS:
            return _error_response("El inventario debe estar en formato Excel (.xlsx).", 400)

        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temporary_file:
                inventory.save(temporary_file)
                temporary_path = Path(temporary_file.name)

            service = InventoryService(storage_directory=Path(app.config["INVENTORY_UPLOADS_DIR"]))
            imported = service.import_inventory(temporary_path)
            set_active_inventory(imported.stored_path, filename, imported.inventory_id)
            payload = inventory_payload()
            if payload.get("inventory", {}).get("is_demonstration"):
                payload["message"] = (
                    f"{filename} se cargó solo como demostración técnica. "
                    "No se usará como evidencia de tesis; carga el kardex real de la tienda."
                )
            else:
                payload["message"] = (
                    f"{filename} se validó y quedó como inventario activo. "
                    "Las cargas anteriores se conservan sin modificarse; confirma con la tienda que este sea el kardex oficial."
                )
            return _json_response(payload)
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
            forecast_dates = [
                date.fromisoformat(item["date"])
                for demands in serialized_predictions.values()
                for item in demands
            ]
            reference_date = min(forecast_dates) - timedelta(days=1)
            product_history, source_dates = _load_product_sale_history(
                catalog_service(), reference_date
            )
            product_forecasts = _build_seasonal_product_forecasts(
                serialized_predictions,
                product_history,
                source_dates,
                products_by_category,
            )
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
                    "product_forecasts_by_category": product_forecasts,
                    "product_allocation_method": (
                        "Participación de ventas de los últimos 120 días anteriores."
                        if source_dates
                        else "Participación histórica disponible."
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

    @app.post("/api/forecast/export")
    def export_forecast():
        """Entrega en Excel la misma vista de compra sugerida del panel."""
        try:
            workbook_bytes, filename = _create_forecast_export(request.get_json(silent=True))
            return send_file(
                workbook_bytes,
                as_attachment=True,
                download_name=filename,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                max_age=0,
            )
        except ValueError as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)

    @app.post("/api/posttest")
    def execute_posttest():
        """Ejecuta una prueba histórica real sin alterar el modelo activo."""
        try:
            body = request.get_json(silent=True) or {}
            cutoff_date = _parse_cutoff_date(body.get("cutoff_date"))
            days = int(body.get("days", 7))
        except (TypeError, ValueError):
            return _error_response("Indica una fecha de corte y un horizonte válidos para el Postest.", 400)

        if cutoff_date is None:
            return _error_response(
                "Selecciona la fecha de corte: debe ser el día anterior al primer día que quieres evaluar.",
                400,
            )
        if not 1 <= days <= 90:
            return _error_response("El Postest admite entre 1 y 90 días.", 400)

        try:
            result = PosttestEvaluationService(CSVDemandRepository()).run(cutoff_date, days)
            period = result["period"]
            try:
                result["inventory"] = inventory_payload(
                    date.fromisoformat(str(period["start_date"])),
                    date.fromisoformat(str(period["end_date"])),
                )
            except (FileNotFoundError, ValueError) as exc:
                # El postest de demanda sigue siendo válido aun cuando el
                # inventario activo no cubra ese período.
                result["inventory"] = {
                    "available": False,
                    "message": f"No se pudieron calcular ISI y TQS para este período: {exc}",
                    "metrics": [],
                }

            record = PosttestResultRepository(Path(app.config["POSTTEST_RESULTS_DIR"])).save(result)
            return _json_response(
                {
                    **result,
                    "run_id": record["run_id"],
                    "created_at": record["created_at"],
                    "message": (
                        "Postest completado: el resultado compara el modelo con las ventas reales "
                        "posteriores y queda guardado para descargarlo."
                    ),
                }
            )
        except (FileNotFoundError, ValueError) as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)

    @app.post("/api/posttest/export")
    def export_posttest():
        """Descarga exactamente el resultado de un Postest guardado."""
        try:
            body = request.get_json(silent=True) or {}
            record = PosttestResultRepository(Path(app.config["POSTTEST_RESULTS_DIR"])).get(
                str(body.get("run_id") or "")
            )
            workbook_bytes, filename = _create_posttest_export(record)
            return send_file(
                workbook_bytes,
                as_attachment=True,
                download_name=filename,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                max_age=0,
            )
        except FileNotFoundError as exc:
            return _error_response(exc, 404)
        except ValueError as exc:
            return _error_response(exc, 400)
        except Exception as exc:
            return _error_response(exc, 500)

    @app.errorhandler(413)
    def file_too_large(_error):
        return _error_response("El archivo supera el límite de 20 MB.", 413)

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


def _load_product_sale_history(
    catalog_service: CatalogService,
    reference_date: date,
) -> tuple[dict[str, dict[str, dict[date, float]]], set[date]]:
    """Lee ventas por SKU anteriores a la fecha que se va a pronosticar.

    La separación por fecha evita usar las ventas reales futuras al ejecutar
    una prueba histórica. Cada SKU queda asociado al nombre canónico y a la
    categoría del catálogo maestro.
    """
    report_paths = sorted(
        {
            *DATA_DIR.glob("Reporte_Taully_*.xlsx"),
            *DATA_DIR.glob("reporte_ventas_*.xlsx"),
        }
    )
    sales_by_product: defaultdict[str, defaultdict[str, defaultdict[date, float]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(float))
    )
    source_dates: set[date] = set()
    sale_reader = ExcelReader()

    for report_path in report_paths:
        sales = sale_reader.read_sales(str(report_path))
        report_dates = {sale.date.date() for sale in sales}
        if len(report_dates) != 1:
            raise ValueError(f"El reporte '{report_path.name}' contiene más de una fecha de venta.")
        report_date = report_dates.pop()
        if report_date > reference_date:
            continue
        if report_date in source_dates:
            raise ValueError(f"Hay más de un reporte para la fecha {report_date:%Y-%m-%d}.")
        source_dates.add(report_date)

        for sale in sales:
            catalog_product = catalog_service.get_product(sale.product_name)
            if catalog_product is None:
                continue
            sales_by_product[catalog_product.category][catalog_product.product_name][report_date] += float(
                sale.quantity
            )

    return (
        {
            category: {product: dict(daily_sales) for product, daily_sales in products.items()}
            for category, products in sales_by_product.items()
        },
        source_dates,
    )


def _build_seasonal_product_forecasts(
    predictions: dict[str, list[dict[str, float | str]]],
    product_history: dict[str, dict[str, dict[date, float]]],
    source_dates: set[date],
    fallback_products_by_category: dict[str, list[dict[str, float | int | str]]],
) -> dict[str, list[dict[str, Any]]]:
    """Distribuye el total de categoría mediante el patrón reciente por SKU.

    El modelo sigue estimando la demanda de la categoría. Para repartirla por
    producto se usan las ventas de los últimos 120 días disponibles. Esta
    ventana es más estable que una muestra corta de un único día de la semana
    y evita sugerir artículos que no tuvieron ventas recientes.
    """
    result: dict[str, list[dict[str, Any]]] = {}
    for category, demands in predictions.items():
        product_series = product_history.get(category, {})
        daily_forecasts: list[dict[str, Any]] = []
        for demand in demands:
            forecast_date = date.fromisoformat(str(demand["date"]))
            weights = _seasonal_product_weights(product_series, source_dates, forecast_date)
            if not weights:
                weights = {
                    str(product["name"]): max(0.0, float(product["historical_share"]))
                    for product in fallback_products_by_category.get(category, [])
                }

            total_weight = sum(weights.values())
            category_quantity = round(float(demand["quantity"]), 2)
            products: list[dict[str, float | str]] = []
            allocated_quantity = 0.0
            if total_weight > 0:
                for product, weight in sorted(weights.items(), key=lambda item: (-item[1], item[0])):
                    allocation_share = weight / total_weight
                    quantity = round(category_quantity * allocation_share, 2)
                    products.append(
                        {
                            "name": product,
                            "quantity": quantity,
                            "allocation_share": round(allocation_share * 100, 2),
                        }
                    )
                    allocated_quantity += quantity

            residual = round(category_quantity - allocated_quantity, 2)
            if residual and products:
                products[0]["quantity"] = round(float(products[0]["quantity"]) + residual, 2)

            daily_forecasts.append(
                {
                    "date": demand["date"],
                    "category_quantity": category_quantity,
                    "products": products,
                }
            )
        result[category] = daily_forecasts
    return result


def _seasonal_product_weights(
    product_series: dict[str, dict[date, float]],
    source_dates: set[date],
    forecast_date: date,
) -> dict[str, float]:
    """Calcula pesos por SKU a partir de las ventas recientes disponibles."""
    recent_dates = [
        sale_date
        for sale_date in source_dates
        if sale_date < forecast_date and (forecast_date - sale_date).days <= 120
    ]
    if not recent_dates:
        return {}

    weights = {
        product: sum(daily_sales.get(sale_date, 0.0) for sale_date in recent_dates)
        for product, daily_sales in product_series.items()
    }
    return {product: weight for product, weight in weights.items() if weight > 0}


def _create_forecast_export(payload: Any) -> tuple[BytesIO, str]:
    """Crea un Excel compacto a partir de la vista actualmente mostrada.

    El navegador envía las filas ya calculadas por el pronóstico. Aquí se
    validan y se vuelve a calcular el redondeo de compra para que el archivo
    conserve el mismo criterio que el panel.
    """
    if not isinstance(payload, dict):
        raise ValueError("No hay datos de pronóstico para exportar.")

    scope = payload.get("scope")
    rows = payload.get("rows")
    if not isinstance(scope, dict) or not isinstance(rows, list) or not rows:
        raise ValueError("No hay filas de pronóstico para exportar.")
    if len(rows) > 2_000:
        raise ValueError("El pronóstico contiene demasiadas filas para exportar.")

    view = scope.get("view")
    if view not in {"period", "date"}:
        raise ValueError("La vista del pronóstico no es válida.")

    try:
        days = int(scope.get("days"))
    except (TypeError, ValueError) as exc:
        raise ValueError("El horizonte del pronóstico no es válido.") from exc
    if not 1 <= days <= 90:
        raise ValueError("El horizonte del pronóstico no es válido.")

    category_scope = str(scope.get("category") or "").strip()

    start_date = _parse_export_date(scope.get("start_date"), "La fecha inicial")
    end_date = _parse_export_date(scope.get("end_date"), "La fecha final")
    if end_date < start_date:
        raise ValueError("El período del pronóstico no es válido.")
    if view == "date" and start_date != end_date:
        raise ValueError("La fecha seleccionada no es válida.")

    normalized_rows: list[dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError("Una fila del pronóstico no es válida.")
        category = str(row.get("category", "")).strip()
        product = str(row.get("product", "")).strip()
        if not category or not product:
            raise ValueError("Cada fila debe incluir categoría y producto.")
        try:
            quantity = float(row.get("quantity"))
            allocation_share = float(row.get("allocation_share", row.get("historical_share", 0)))
        except (TypeError, ValueError) as exc:
            raise ValueError("La cantidad estimada no es válida.") from exc
        if not math.isfinite(quantity) or quantity < 0 or not math.isfinite(allocation_share):
            raise ValueError("La cantidad estimada no es válida.")

        suggested_packages = row.get("suggested_packages")
        if suggested_packages is None:
            suggested_packages = math.ceil(quantity)
        try:
            suggested_value = float(suggested_packages)
        except (TypeError, ValueError) as exc:
            raise ValueError("La compra sugerida no es válida.") from exc
        if not math.isfinite(suggested_value) or suggested_value < 0 or not suggested_value.is_integer():
            raise ValueError("La compra sugerida no es válida.")
        suggested_packages = int(suggested_value)

        normalized_rows.append(
            {
                "category": category,
                "product": product,
                "quantity": round(quantity, 2),
                "suggested_packages": suggested_packages,
                "allocation_share": max(0.0, allocation_share) / 100,
            }
        )

    if category_scope and any(row["category"] != category_scope for row in normalized_rows):
        raise ValueError("El archivo por categoría solo puede incluir productos de esa categoría.")

    period_label = (
        start_date.strftime("%d/%m/%Y")
        if start_date == end_date
        else f"Del {start_date.strftime('%d/%m/%Y')} al {end_date.strftime('%d/%m/%Y')}"
    )
    mode_label = "Prueba histórica" if scope.get("cutoff_date") else "Pronóstico futuro"
    category_label = category_scope or "Todas las categorías"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Pronóstico"
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "A6"
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0

    navy = "091F35"
    gold = "F0C400"
    line = "D9E2E8"
    muted = "64748B"
    thin_line = Side(style="thin", color=line)

    worksheet["A1"] = f"Pronóstico de compra sugerida - {category_label}"
    worksheet["A1"].font = Font(name="Arial", size=14, bold=True, color=navy)
    worksheet["A1"].border = Border(bottom=Side(style="medium", color=gold))

    metadata = [
        ("Tipo de consulta", mode_label),
        ("Período consultado", period_label),
        ("Horizonte (días)", days),
        ("Criterio", "Mezcla de ventas recientes y paquetes enteros asignados por categoría."),
    ]
    for row_number, (label, value) in enumerate(metadata, start=2):
        worksheet.cell(row=row_number, column=1, value=label)
        worksheet.cell(row=row_number, column=1).font = Font(name="Arial", size=10, bold=True, color=navy)
        worksheet.cell(row=row_number, column=2, value=value)
        worksheet.cell(row=row_number, column=2).font = Font(name="Arial", size=10, color=muted)

    headers = [
        "CATEGORÍA",
        "PRODUCTO",
        "DEMANDA ESTIMADA",
        "COMPRA SUGERIDA",
        "PARTICIPACIÓN DE ASIGNACIÓN",
    ]
    for column, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=6, column=column, value=header)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin_line, right=Side(style="thin", color="FFFFFF"))

    for row_number, item in enumerate(normalized_rows, start=7):
        values = [
            _excel_safe_text(item["category"]),
            _excel_safe_text(item["product"]),
            item["quantity"],
            item["suggested_packages"],
            item["allocation_share"],
        ]
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row_number, column=column, value=value)
            cell.font = Font(name="Arial", size=10, color=navy)
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=thin_line)
        worksheet.cell(row=row_number, column=3).number_format = "#,##0.00"
        worksheet.cell(row=row_number, column=4).number_format = "#,##0"
        worksheet.cell(row=row_number, column=5).number_format = "0.0%"

    worksheet.auto_filter.ref = f"A6:E{len(normalized_rows) + 6}"
    worksheet.column_dimensions["A"].width = 20
    worksheet.column_dimensions["B"].width = 52
    worksheet.column_dimensions["C"].width = 20
    worksheet.column_dimensions["D"].width = 20
    worksheet.column_dimensions["E"].width = 25
    worksheet.row_dimensions[1].height = 24
    worksheet.row_dimensions[6].height = 22

    filename_period = start_date.isoformat() if start_date == end_date else f"{start_date.isoformat()}_a_{end_date.isoformat()}"
    category_filename = f"{secure_filename(category_scope.lower())}_" if category_scope else ""
    workbook_bytes = BytesIO()
    workbook.save(workbook_bytes)
    workbook_bytes.seek(0)
    return workbook_bytes, f"pronostico_compra_{category_filename}{filename_period}.xlsx"


def _create_posttest_export(record: Any) -> tuple[BytesIO, str]:
    """Crea evidencia Excel desde un resultado de Postest ya almacenado.

    El navegador solo entrega el identificador. Así el archivo descargado no
    depende de filas editables en la pantalla y conserva el mismo resultado
    que calculó el servicio para las ventas reales del período.
    """
    if not isinstance(record, dict) or not isinstance(record.get("result"), dict):
        raise ValueError("No se encontró un resultado válido de Postest para exportar.")

    result = record["result"]
    model = result.get("model") if isinstance(result.get("model"), dict) else {}
    baseline = result.get("baseline") if isinstance(result.get("baseline"), dict) else {}
    model_metrics = model.get("metrics") if isinstance(model.get("metrics"), dict) else {}
    baseline_metrics = baseline.get("metrics") if isinstance(baseline.get("metrics"), dict) else {}
    period = result.get("period") if isinstance(result.get("period"), dict) else {}
    observations = result.get("observations") if isinstance(result.get("observations"), list) else []
    inventory = result.get("inventory") if isinstance(result.get("inventory"), dict) else {}
    inventory_info = inventory.get("inventory") if isinstance(inventory.get("inventory"), dict) else {}
    training = result.get("training") if isinstance(result.get("training"), dict) else {}
    methodology = result.get("methodology") if isinstance(result.get("methodology"), dict) else {}

    try:
        cutoff_date = _parse_export_date(result.get("cutoff_date"), "La fecha de corte")
        period_start = _parse_export_date(period.get("start_date"), "La fecha inicial")
        period_end = _parse_export_date(period.get("end_date"), "La fecha final")
        horizon_days = int(result.get("horizon_days"))
    except (TypeError, ValueError) as exc:
        raise ValueError("El resultado guardado del Postest no tiene un período válido.") from exc
    if not 1 <= horizon_days <= 90 or period_end < period_start:
        raise ValueError("El resultado guardado del Postest no tiene un período válido.")

    navy = "091F35"
    gold = "F0C400"
    blue = "0B86B5"
    line = "D9E2E8"
    muted = "64748B"
    thin_line = Side(style="thin", color=line)

    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "Resumen postest"
    summary_sheet.sheet_view.showGridLines = False
    summary_sheet.sheet_properties.pageSetUpPr.fitToPage = True
    summary_sheet.page_setup.orientation = "landscape"
    summary_sheet.page_setup.fitToWidth = 1
    summary_sheet.page_setup.fitToHeight = 0

    summary_sheet.merge_cells("A1:E1")
    title = summary_sheet["A1"]
    title.value = "Postest de pronóstico y comparación con PMS-7"
    title.font = Font(name="Arial", size=15, bold=True, color=navy)
    title.alignment = Alignment(vertical="center")
    title.fill = PatternFill("solid", fgColor="F7FBFD")
    title.border = Border(bottom=Side(style="medium", color=gold))
    summary_sheet.row_dimensions[1].height = 27

    evaluated_categories = result.get("categories_evaluated")
    evaluated_names = [str(category) for category in evaluated_categories] if isinstance(evaluated_categories, list) else []
    excluded_categories = training.get("excluded_categories")
    if isinstance(excluded_categories, dict):
        excluded_names = [
            f"{category}: {reason}" for category, reason in sorted(excluded_categories.items())
        ]
    elif isinstance(excluded_categories, list):
        excluded_names = [str(category) for category in excluded_categories]
    else:
        excluded_names = []

    inventory_source = "No se cargó un inventario aplicable al período."
    if inventory.get("available"):
        inventory_source = (
            f"{inventory_info.get('name') or 'Inventario activo'} · "
            f"{inventory_info.get('source_kind') or 'origen no indicado'} · "
            f"ID SHA-256: {inventory_info.get('source_id') or 'no disponible'}"
        )
    inventory_warnings = [
        str(warning) for warning in inventory.get("warnings", []) if str(warning).strip()
    ] if isinstance(inventory.get("warnings"), list) else []

    metadata = [
        ("Corte de entrenamiento", cutoff_date.strftime("%d/%m/%Y")),
        ("Período evaluado", f"Del {period_start.strftime('%d/%m/%Y')} al {period_end.strftime('%d/%m/%Y')}"),
        ("Horizonte", f"{horizon_days} días"),
        ("Ventas reales", "Historial consolidado posterior al corte."),
        ("Resultado guardado", str(record.get("created_at") or "—")),
        ("Huella del historial (SHA-256)", str(methodology.get("history_fingerprint") or "No disponible")),
        ("Categorías evaluadas", ", ".join(evaluated_names) or "No disponible"),
        (
            "Alcance de métricas globales",
            "Las métricas globales consideran únicamente las categorías evaluadas por el modelo.",
        ),
        ("Categorías excluidas", "; ".join(excluded_names) if excluded_names else "Ninguna"),
        ("Inventario del período", inventory_source),
    ]
    if inventory_warnings:
        metadata.append(("Advertencia de inventario", " ".join(inventory_warnings)))
    for row_number, (label, value) in enumerate(metadata, start=2):
        label_cell = summary_sheet.cell(row=row_number, column=1, value=label)
        label_cell.font = Font(name="Arial", size=10, bold=True, color=navy)
        value_cell = summary_sheet.cell(
            row=row_number, column=2, value=_excel_safe_text(str(value))
        )
        value_cell.font = Font(name="Arial", size=10, color=muted)
        value_cell.alignment = Alignment(vertical="top", wrap_text=True)
        summary_sheet.merge_cells(start_row=row_number, start_column=2, end_row=row_number, end_column=5)
        if len(str(value)) > 120:
            summary_sheet.row_dimensions[row_number].height = 32

    metric_header_row = len(metadata) + 3
    summary_sheet.freeze_panes = f"A{metric_header_row}"
    metric_headers = ["INDICADOR", "MODELO PREDICTIVO", "PMS-7", "MEJORA DEL MODELO", "MEJOR RESULTADO"]
    for column, header in enumerate(metric_headers, start=1):
        cell = summary_sheet.cell(row=metric_header_row, column=column, value=header)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin_line, right=Side(style="thin", color="FFFFFF"))

    comparison = result.get("comparison") if isinstance(result.get("comparison"), dict) else {}
    improvements = comparison.get("improvement_vs_baseline") if isinstance(comparison.get("improvement_vs_baseline"), dict) else {}
    winners = comparison.get("winner_by_metric") if isinstance(comparison.get("winner_by_metric"), dict) else {}
    metric_names = {"mae": "MAE (unidades)", "rmse": "RMSE (unidades)", "mape": "MAPE", "wape": "WAPE"}
    for row_number, metric_key in enumerate(("mae", "rmse", "mape", "wape"), start=metric_header_row + 1):
        summary_sheet.cell(row=row_number, column=1, value=metric_names[metric_key])
        summary_sheet.cell(row=row_number, column=2, value=_excel_number(model_metrics.get(metric_key)))
        summary_sheet.cell(row=row_number, column=3, value=_excel_number(baseline_metrics.get(metric_key)))
        improvement = improvements.get(metric_key) if isinstance(improvements.get(metric_key), dict) else {}
        summary_sheet.cell(row=row_number, column=4, value=_excel_number(improvement.get("percent")))
        winner = str(winners.get(metric_key) or "not_available")
        winner_label = {"model": "Modelo predictivo", "pms_7": "PMS-7", "tie": "Empate"}.get(winner, "No disponible")
        summary_sheet.cell(row=row_number, column=5, value=winner_label)
        for column in range(1, 6):
            cell = summary_sheet.cell(row=row_number, column=column)
            cell.font = Font(name="Arial", size=10, color=navy)
            cell.border = Border(bottom=thin_line)
            cell.alignment = Alignment(vertical="center", horizontal="center" if column > 1 else "left")
        for column in (2, 3):
            summary_sheet.cell(row=row_number, column=column).number_format = '0.00"%"' if metric_key in {"mape", "wape"} else "#,##0.00"
        summary_sheet.cell(row=row_number, column=4).number_format = '0.00"%"'

    category_start = metric_header_row + 8
    summary_sheet.merge_cells(start_row=category_start, start_column=1, end_row=category_start, end_column=5)
    category_title = summary_sheet.cell(row=category_start, column=1, value="Resultados por categoría")
    category_title.font = Font(name="Arial", size=11, bold=True, color=navy)
    category_title.fill = PatternFill("solid", fgColor="EAF7FB")
    category_headers = ["CATEGORÍA", "DEMANDA REAL", "MODELO (WAPE)", "PMS-7 (WAPE)", "MEJOR WAPE"]
    for column, header in enumerate(category_headers, start=1):
        cell = summary_sheet.cell(row=category_start + 1, column=column, value=header)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    model_by_category = model.get("by_category") if isinstance(model.get("by_category"), dict) else {}
    baseline_by_category = baseline.get("by_category") if isinstance(baseline.get("by_category"), dict) else {}
    categories = sorted(set(model_by_category) | set(baseline_by_category))
    for row_number, category in enumerate(categories, start=category_start + 2):
        model_category = model_by_category.get(category) if isinstance(model_by_category.get(category), dict) else {}
        baseline_category = baseline_by_category.get(category) if isinstance(baseline_by_category.get(category), dict) else {}
        model_wape = _excel_number(model_category.get("wape"))
        baseline_wape = _excel_number(baseline_category.get("wape"))
        if model_wape is None or baseline_wape is None:
            winner = "No disponible"
        elif math.isclose(model_wape, baseline_wape, rel_tol=0, abs_tol=1e-9):
            winner = "Empate"
        elif model_wape < baseline_wape:
            winner = "Modelo predictivo"
        else:
            winner = "PMS-7"
        values = [
            _excel_safe_text(str(category)),
            _excel_number(model_category.get("actual_total")),
            model_wape,
            baseline_wape,
            winner,
        ]
        for column, value in enumerate(values, start=1):
            cell = summary_sheet.cell(row=row_number, column=column, value=value)
            cell.font = Font(name="Arial", size=10, color=navy)
            cell.border = Border(bottom=thin_line)
        summary_sheet.cell(row=row_number, column=2).number_format = "#,##0.00"
        summary_sheet.cell(row=row_number, column=3).number_format = '0.00"%"'
        summary_sheet.cell(row=row_number, column=4).number_format = '0.00"%"'

    summary_sheet.column_dimensions["A"].width = 26
    summary_sheet.column_dimensions["B"].width = 24
    summary_sheet.column_dimensions["C"].width = 22
    summary_sheet.column_dimensions["D"].width = 22
    summary_sheet.column_dimensions["E"].width = 24

    # La pantalla mantiene la comparación por WAPE compacta, pero el archivo
    # de evidencia conserva los cuatro indicadores por categoría para que el
    # análisis pueda revisarse sin recalcular nada fuera del programa.
    category_metrics_sheet = workbook.create_sheet("Métricas categoría")
    category_metrics_sheet.sheet_view.showGridLines = False
    category_metrics_sheet.freeze_panes = "A2"
    category_metric_headers = [
        "CATEGORÍA", "DEMANDA REAL", "MODELO MAE", "MODELO RMSE",
        "MODELO MAPE", "MODELO WAPE", "PMS-7 MAE", "PMS-7 RMSE",
        "PMS-7 MAPE", "PMS-7 WAPE", "MEJOR WAPE",
    ]
    for column, header in enumerate(category_metric_headers, start=1):
        cell = category_metrics_sheet.cell(row=1, column=column, value=header)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin_line, right=Side(style="thin", color="FFFFFF"))

    for row_number, category in enumerate(categories, start=2):
        model_category = model_by_category.get(category) if isinstance(model_by_category.get(category), dict) else {}
        baseline_category = baseline_by_category.get(category) if isinstance(baseline_by_category.get(category), dict) else {}
        model_wape = _excel_number(model_category.get("wape"))
        baseline_wape = _excel_number(baseline_category.get("wape"))
        if model_wape is None or baseline_wape is None:
            winner = "No disponible"
        elif math.isclose(model_wape, baseline_wape, rel_tol=0, abs_tol=1e-9):
            winner = "Empate"
        elif model_wape < baseline_wape:
            winner = "Modelo predictivo"
        else:
            winner = "PMS-7"
        values = [
            _excel_safe_text(str(category)),
            _excel_number(model_category.get("actual_total", baseline_category.get("actual_total"))),
            _excel_number(model_category.get("mae")),
            _excel_number(model_category.get("rmse")),
            _excel_number(model_category.get("mape")),
            model_wape,
            _excel_number(baseline_category.get("mae")),
            _excel_number(baseline_category.get("rmse")),
            _excel_number(baseline_category.get("mape")),
            baseline_wape,
            winner,
        ]
        for column, value in enumerate(values, start=1):
            cell = category_metrics_sheet.cell(row=row_number, column=column, value=value)
            cell.font = Font(name="Arial", size=10, color=navy)
            cell.border = Border(bottom=thin_line)
            if column in {5, 6, 9, 10}:
                cell.number_format = '0.00"%"'
            elif 2 <= column <= 4 or 7 <= column <= 8:
                cell.number_format = "#,##0.00"
    for column, width in {
        "A": 24, "B": 18, "C": 15, "D": 16, "E": 16, "F": 16,
        "G": 15, "H": 16, "I": 16, "J": 16, "K": 21,
    }.items():
        category_metrics_sheet.column_dimensions[column].width = width
    category_metrics_sheet.auto_filter.ref = f"A1:K{max(1, len(categories) + 1)}"

    detail_sheet = workbook.create_sheet("Detalle diario")
    detail_sheet.sheet_view.showGridLines = False
    detail_sheet.freeze_panes = "A2"
    detail_headers = ["FECHA", "CATEGORÍA", "DEMANDA REAL", "MODELO PREDICTIVO", "PMS-7", "ERROR ABS. MODELO", "ERROR ABS. PMS-7"]
    for column, header in enumerate(detail_headers, start=1):
        cell = detail_sheet.cell(row=1, column=column, value=header)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row_number, observation in enumerate(observations, start=2):
        if not isinstance(observation, dict):
            continue
        actual = _excel_number(observation.get("actual_demand"))
        prediction = _excel_number(observation.get("model_prediction"))
        baseline_prediction = _excel_number(observation.get("pms_7_prediction"))
        values = [
            str(observation.get("date") or ""),
            _excel_safe_text(str(observation.get("category") or "")),
            actual,
            prediction,
            baseline_prediction,
            abs(actual - prediction) if actual is not None and prediction is not None else None,
            abs(actual - baseline_prediction) if actual is not None and baseline_prediction is not None else None,
        ]
        for column, value in enumerate(values, start=1):
            cell = detail_sheet.cell(row=row_number, column=column, value=value)
            cell.font = Font(name="Arial", size=10, color=navy)
            cell.border = Border(bottom=thin_line)
            if column >= 3:
                cell.number_format = "#,##0.00"
    for column, width in {"A": 16, "B": 23, "C": 18, "D": 22, "E": 14, "F": 21, "G": 20}.items():
        detail_sheet.column_dimensions[column].width = width
    detail_sheet.auto_filter.ref = f"A1:G{max(1, len(observations) + 1)}"

    if inventory.get("available") and isinstance(inventory.get("metrics"), list):
        inventory_sheet = workbook.create_sheet("Inventario")
        inventory_sheet.sheet_view.showGridLines = False
        inventory_sheet["A1"] = "Indicadores de inventario del período del Postest"
        inventory_sheet["A1"].font = Font(name="Arial", size=14, bold=True, color=navy)
        inventory_sheet["A2"] = "ISI = CD / (SI + EN). TQS = DQS / DD. Los resultados dependen del archivo de inventario activo."
        inventory_sheet["A2"].font = Font(name="Arial", size=9, italic=True, color=muted)
        inventory_sheet.merge_cells("A2:K2")
        inventory_sheet["A3"] = _excel_safe_text(inventory_source)
        inventory_sheet["A3"].font = Font(name="Arial", size=9, color=muted)
        inventory_sheet["A3"].alignment = Alignment(vertical="top", wrap_text=True)
        inventory_sheet.merge_cells("A3:K3")
        if inventory_warnings:
            inventory_sheet["A4"] = _excel_safe_text("ADVERTENCIA: " + " ".join(inventory_warnings))
            inventory_sheet["A4"].font = Font(name="Arial", size=9, bold=True, color="9B1C1C")
            inventory_sheet["A4"].alignment = Alignment(vertical="top", wrap_text=True)
            inventory_sheet.merge_cells("A4:K4")
            inventory_sheet.row_dimensions[4].height = 34
        inventory_header_row = 6
        inventory_headers = [
            "CATEGORÍA", "SI", "EN", "SF", "CD", "ISI", "DQS", "DD",
            "DD REGISTRADOS", "COBERTURA", "TQS",
        ]
        for column, header in enumerate(inventory_headers, start=1):
            cell = inventory_sheet.cell(row=inventory_header_row, column=column, value=header)
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row_number, metric in enumerate(inventory["metrics"], start=inventory_header_row + 1):
            if not isinstance(metric, dict):
                continue
            values = [
                _excel_safe_text(str(metric.get("category") or "")),
                _excel_number(metric.get("SI")), _excel_number(metric.get("EN")),
                _excel_number(metric.get("SF")), _excel_number(metric.get("CD")),
                _excel_number(metric.get("ISI")), metric.get("DQS"), metric.get("DD"),
                metric.get("DD_REGISTRADOS"), _excel_number(metric.get("COBERTURA")),
                _excel_number(metric.get("TQS")),
            ]
            for column, value in enumerate(values, start=1):
                cell = inventory_sheet.cell(row=row_number, column=column, value=value)
                cell.font = Font(name="Arial", size=10, color=navy)
                cell.border = Border(bottom=thin_line)
                if column in {6, 10, 11}:
                    cell.number_format = '0.00"%"'
                elif column in {2, 3, 4, 5}:
                    cell.number_format = "#,##0.00"
        for column, width in {
            "A": 22, "B": 14, "C": 14, "D": 14, "E": 14, "F": 12,
            "G": 11, "H": 11, "I": 17, "J": 14, "K": 12,
        }.items():
            inventory_sheet.column_dimensions[column].width = width
        inventory_sheet.auto_filter.ref = f"A{inventory_header_row}:K{max(inventory_header_row, len(inventory['metrics']) + inventory_header_row)}"

    workbook_bytes = BytesIO()
    workbook.save(workbook_bytes)
    workbook_bytes.seek(0)
    return workbook_bytes, f"postest_pronostico_{period_start.isoformat()}_a_{period_end.isoformat()}.xlsx"


def _excel_number(value: Any) -> float | None:
    """Convierte a número seguro para escribir en Excel, sin fórmulas ni NaN."""
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _parse_export_date(value: Any, message: str) -> date:
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{message} no es válida.") from exc


def _excel_safe_text(value: str) -> str:
    """Evita que nombres de productos se interpreten como fórmulas en Excel."""
    return f"'{value}" if value[:1] in {"=", "+", "-", "@"} else value
