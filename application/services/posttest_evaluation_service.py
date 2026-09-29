"""Evaluación histórica reproducible del pronóstico (postest).

Este servicio vuelve a entrenar el modelo con las ventas conocidas hasta una
fecha de corte y contrasta las predicciones de los días posteriores con las
ventas reales ya registradas.  El artefacto del modelo se crea en un directorio
temporal: nunca reemplaza el modelo que el usuario entrena desde el panel.
"""

from __future__ import annotations

from contextlib import contextmanager
from datetime import date, datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Iterator

import numpy as np

from application.services.predict_service import PredictService
from application.services.train_service import TrainService
from application.services.history_fingerprint import history_fingerprint
from domain.entities.demand import Demand
from domain.interfaces.repositories import DemandRepository


class PosttestEvaluationService:
    """Calcula un postest de pronóstico con el modelo actual y PMS-7.

    El resultado contiene únicamente tipos serializables por JSON (dict,
    list, str, int, float, bool o ``None``), por lo que puede usarse tal cual
    en una ruta web o en una exportación Excel.
    """

    def __init__(self, demand_repo: DemandRepository):
        self.demand_repo = demand_repo

    def run(self, cutoff_date: date | datetime, horizon_days: int = 7) -> dict[str, Any]:
        """Ejecuta el postest para ``horizon_days`` posteriores al corte.

        Args:
            cutoff_date: último día de ventas disponible para entrenar.
            horizon_days: cantidad de días consecutivos a evaluar.

        Raises:
            ValueError: si el corte, horizonte o ventas reales posteriores no
                permiten una comparación histórica completa.
        """
        cutoff = self._normalise_cutoff(cutoff_date)
        horizon = self._normalise_horizon(horizon_days)
        all_demands = list(self.demand_repo.get_all_demands())
        if not all_demands:
            raise ValueError("No hay ventas históricas para ejecutar el postest.")

        historical_dates = {demand.date.date() for demand in all_demands}
        if cutoff not in historical_dates:
            raise ValueError(
                "La fecha de corte debe corresponder a un día de ventas procesado."
            )

        evaluation_dates = [cutoff + timedelta(days=offset) for offset in range(1, horizon + 1)]
        missing_dates = [
            evaluation_day.isoformat()
            for evaluation_day in evaluation_dates
            if evaluation_day not in historical_dates
        ]
        if missing_dates:
            raise ValueError(
                "No hay ventas reales completas después de la fecha de corte para el postest: "
                + ", ".join(missing_dates)
            )

        # Se usa exactamente TrainService/PredictService, pero con un archivo
        # temporal exclusivo. Así el modelo que el usuario entrenó desde el
        # panel no se reemplaza ni se comparte con esta validación.
        with self._temporary_model_artifact() as temporary_model_path:
            train_service = TrainService(self.demand_repo, models_file=temporary_model_path)
            training_metrics = train_service.run(cutoff_date=cutoff)
            predictions = PredictService(self.demand_repo, models_file=temporary_model_path).predict_future(
                days=horizon,
                cutoff_date=cutoff,
            )

            # El archivo existe durante la ejecución para que sea evidente que
            # el servicio no está usando ni modificando el modelo del usuario.
            if not temporary_model_path.exists():
                raise RuntimeError("No se pudo crear el artefacto temporal del postest.")

        if not predictions:
            raise ValueError("El entrenamiento no produjo categorías para evaluar en el postest.")

        categories = sorted(predictions)
        actual_lookup = self._actual_lookup(all_demands)
        missing_actuals = [
            f"{category} ({evaluation_day.isoformat()})"
            for category in categories
            for evaluation_day in evaluation_dates
            if (evaluation_day, category) not in actual_lookup
        ]
        if missing_actuals:
            preview = ", ".join(missing_actuals[:5])
            suffix = "..." if len(missing_actuals) > 5 else ""
            raise ValueError(
                "No hay ventas reales completas después de la fecha de corte para el postest: "
                f"{preview}{suffix}"
            )

        observations = self._build_observations(
            all_demands=all_demands,
            predictions=predictions,
            actual_lookup=actual_lookup,
            cutoff=cutoff,
            evaluation_dates=evaluation_dates,
        )
        model_metrics = self._metrics_for(observations, "model_prediction")
        baseline_metrics = self._metrics_for(observations, "pms_7_prediction")
        training_demands = [
            demand for demand in all_demands if demand.date.date() <= cutoff
        ]

        return {
            "cutoff_date": cutoff.isoformat(),
            "horizon_days": horizon,
            "period": {
                "start_date": evaluation_dates[0].isoformat(),
                "end_date": evaluation_dates[-1].isoformat(),
            },
            "categories_evaluated": categories,
            "observation_count": len(observations),
            "training": {
                "internal_validation_metrics": self._serialise_metrics(training_metrics),
                "trained_categories": int(train_service.last_training_summary["trained_categories"]),
                "excluded_categories": dict(train_service.last_training_summary["excluded_categories"]),
                "validation_by_category": dict(train_service.last_training_summary["validation_by_category"]),
                "temporary_model_artifact": True,
            },
            "model": {
                "name": "Modelo predictivo",
                "metrics": model_metrics,
                "by_category": self._metrics_by_category(observations, "model_prediction"),
            },
            "baseline": {
                "name": "PMS-7",
                "description": "Promedio móvil simple de 7 días, calculado de forma recursiva.",
                "metrics": baseline_metrics,
                "by_category": self._metrics_by_category(observations, "pms_7_prediction"),
            },
            "comparison": self._comparison(model_metrics, baseline_metrics),
            "observations": observations,
            "methodology": {
                "model": (
                    "El modelo se entrena solo con las ventas hasta la fecha de corte "
                    "en un artefacto temporal y pronostica los días posteriores."
                ),
                "actuals": "Las ventas reales posteriores se toman del historial procesado.",
                "baseline": "PMS-7 usa las últimas siete observaciones y añade cada estimación al siguiente paso.",
                # Las huellas permiten comprobar posteriormente que una
                # exportación corresponde al mismo historial de ventas, aun si
                # el archivo CSV se corrige o se amplía después del Postest.
                "history_fingerprint": history_fingerprint(all_demands),
                "training_history_fingerprint": history_fingerprint(training_demands),
                "history_records": len(all_demands),
                "training_history_records": len(training_demands),
            },
        }

    # Nombre alternativo útil para controladores que expresan la acción como
    # evaluación en vez de ejecución.
    evaluate = run

    @staticmethod
    def _normalise_cutoff(value: date | datetime) -> date:
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        raise ValueError("La fecha de corte debe ser una fecha válida.")

    @staticmethod
    def _normalise_horizon(value: int) -> int:
        if isinstance(value, bool) or not isinstance(value, int) or value < 1:
            raise ValueError("El horizonte del postest debe ser un número entero mayor que cero.")
        return value

    @staticmethod
    @contextmanager
    def _temporary_model_artifact() -> Iterator[Path]:
        """Entrega una ruta temporal exclusiva y la elimina al finalizar."""
        with TemporaryDirectory(prefix="taully_posttest_") as temporary_directory:
            yield Path(temporary_directory) / "posttest_models_by_category.pkl"

    @staticmethod
    def _actual_lookup(all_demands: list[Demand]) -> dict[tuple[date, str], float]:
        """Consolida defensivamente una venta por fecha y categoría."""
        lookup: dict[tuple[date, str], float] = {}
        for demand in all_demands:
            key = (demand.date.date(), str(demand.category))
            lookup[key] = lookup.get(key, 0.0) + float(demand.quantity)
        return lookup

    def _build_observations(
        self,
        *,
        all_demands: list[Demand],
        predictions: dict[str, list[Demand]],
        actual_lookup: dict[tuple[date, str], float],
        cutoff: date,
        evaluation_dates: list[date],
    ) -> list[dict[str, Any]]:
        observations: list[dict[str, Any]] = []
        for category in sorted(predictions):
            model_predictions = predictions[category]
            if len(model_predictions) != len(evaluation_dates):
                raise ValueError(
                    f"El modelo no generó las {len(evaluation_dates)} fechas esperadas para {category}."
                )

            expected_prediction_dates = [prediction.date.date() for prediction in model_predictions]
            if expected_prediction_dates != evaluation_dates:
                raise ValueError(
                    f"Las fechas pronosticadas para {category} no coinciden con el período del postest."
                )

            pms_7_predictions = self._recursive_pms_7(
                all_demands=all_demands,
                category=category,
                cutoff=cutoff,
                horizon=len(evaluation_dates),
            )
            for evaluation_day, model_prediction, pms_prediction in zip(
                evaluation_dates,
                model_predictions,
                pms_7_predictions,
            ):
                observations.append(
                    {
                        "date": evaluation_day.isoformat(),
                        "category": category,
                        "actual_demand": self._number(actual_lookup[(evaluation_day, category)]),
                        "model_prediction": self._number(float(model_prediction.quantity)),
                        "pms_7_prediction": self._number(pms_prediction),
                    }
                )
        return observations

    @staticmethod
    def _recursive_pms_7(
        *,
        all_demands: list[Demand],
        category: str,
        cutoff: date,
        horizon: int,
    ) -> list[float]:
        historical_values = [
            float(demand.quantity)
            for demand in sorted(all_demands, key=lambda item: item.date)
            if str(demand.category) == category and demand.date.date() <= cutoff
        ]
        if len(historical_values) < 7:
            raise ValueError(f"No hay siete observaciones históricas para calcular PMS-7 en {category}.")

        predictions: list[float] = []
        values = list(historical_values)
        for _ in range(horizon):
            prediction = float(np.mean(values[-7:]))
            predictions.append(prediction)
            values.append(prediction)
        return predictions

    @classmethod
    def _metrics_for(cls, observations: list[dict[str, Any]], prediction_field: str) -> dict[str, float | None]:
        actual = np.asarray([row["actual_demand"] for row in observations], dtype=float)
        predicted = np.asarray([row[prediction_field] for row in observations], dtype=float)
        errors = actual - predicted
        absolute_errors = np.abs(errors)
        nonzero_actuals = actual != 0
        actual_total = float(np.abs(actual).sum())

        mape: float | None
        if np.any(nonzero_actuals):
            mape = float(np.mean(absolute_errors[nonzero_actuals] / np.abs(actual[nonzero_actuals])) * 100)
        else:
            mape = None

        wape = float(absolute_errors.sum() / actual_total * 100) if actual_total else None
        return {
            "mae": cls._number(float(np.mean(absolute_errors))),
            "rmse": cls._number(float(np.sqrt(np.mean(np.square(errors))))),
            "mape": cls._number(mape) if mape is not None else None,
            "wape": cls._number(wape) if wape is not None else None,
            "actual_total": cls._number(float(actual.sum())),
            "prediction_total": cls._number(float(predicted.sum())),
        }

    @classmethod
    def _metrics_by_category(
        cls,
        observations: list[dict[str, Any]],
        prediction_field: str,
    ) -> dict[str, dict[str, float | None]]:
        return {
            category: cls._metrics_for(
                [row for row in observations if row["category"] == category],
                prediction_field,
            )
            for category in sorted({str(row["category"]) for row in observations})
        }

    @classmethod
    def _comparison(
        cls,
        model_metrics: dict[str, float | None],
        baseline_metrics: dict[str, float | None],
    ) -> dict[str, Any]:
        improvement: dict[str, dict[str, float | None]] = {}
        winner: dict[str, str] = {}
        for metric in ("mae", "rmse", "mape", "wape"):
            model_value = model_metrics[metric]
            baseline_value = baseline_metrics[metric]
            if model_value is None or baseline_value is None:
                improvement[metric] = {"absolute": None, "percent": None}
                winner[metric] = "not_available"
                continue

            absolute = float(baseline_value - model_value)
            relative = float(absolute / baseline_value * 100) if baseline_value else None
            improvement[metric] = {
                "absolute": cls._number(absolute),
                "percent": cls._number(relative) if relative is not None else None,
            }
            if np.isclose(model_value, baseline_value):
                winner[metric] = "tie"
            elif model_value < baseline_value:
                winner[metric] = "model"
            else:
                winner[metric] = "pms_7"

        return {
            "baseline_name": "PMS-7",
            "candidate_name": "Modelo predictivo",
            "improvement_vs_baseline": improvement,
            "winner_by_metric": winner,
            "interpretation": "Un valor positivo de mejora indica que el modelo redujo el error frente a PMS-7.",
        }

    @classmethod
    def _serialise_metrics(cls, metrics: dict[str, float]) -> dict[str, float | None]:
        return {str(name): cls._number(float(value)) for name, value in metrics.items()}

    @staticmethod
    def _number(value: float | None) -> float:
        """Convierte valores NumPy a float JSON y limita ruido de precisión."""
        if value is None:
            return None  # type: ignore[return-value]
        if not np.isfinite(value):
            raise ValueError("El postest produjo un indicador no finito.")
        return round(float(value), 6)
