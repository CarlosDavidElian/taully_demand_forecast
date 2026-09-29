"""Importación y cálculo de indicadores de inventario por categoría."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any
import unicodedata
from zipfile import BadZipFile

import numpy as np
import pandas as pd

from infrastructure.repositories.inventory_repository import InventoryWorkbookStore, StoredInventoryFile


@dataclass(frozen=True)
class InventoryImportResult:
    """Resultado de validar y conservar un archivo de inventario."""

    inventory_id: str | None
    stored_path: Path
    original_name: str
    records: int
    categories: int
    start_date: date
    end_date: date
    reconciliation_rows: int = 0

    def to_dict(self) -> dict[str, Any]:
        result = asdict(self)
        result["stored_path"] = str(self.stored_path)
        result["start_date"] = self.start_date.isoformat()
        result["end_date"] = self.end_date.isoformat()
        return result


@dataclass(frozen=True)
class InventoryCategoryMetrics:
    """Indicadores calculados para una categoría y periodo evaluado.

    ``CD`` se obtiene por el balance solicitado: ``SI + EN - SF``. Para no
    sumar saldos diarios, ``SI`` corresponde al primer saldo de cada producto
    dentro del período y ``SF`` al último. ``DQS`` cuenta fechas únicas con al
    menos una alerta positiva de falta de stock; ``DD`` son los días calendario
    del período seleccionado.
    """

    category: str
    start_date: date
    end_date: date
    records: int
    stock_initial: float
    entries: float
    stock_final: float
    demand_quantity: float
    inventory_output_index: float
    stockout_days: int
    evaluated_days: int
    registered_days: int
    coverage_rate: float
    stockout_rate: float

    def to_dict(self) -> dict[str, Any]:
        return {
            "category": self.category,
            "start_date": self.start_date.isoformat(),
            "end_date": self.end_date.isoformat(),
            "records": self.records,
            "SI": self.stock_initial,
            "EN": self.entries,
            "SF": self.stock_final,
            "CD": self.demand_quantity,
            "ISI": self.inventory_output_index,
            "DQS": self.stockout_days,
            "DD": self.evaluated_days,
            "DD_REGISTRADOS": self.registered_days,
            "COBERTURA": self.coverage_rate,
            "TQS": self.stockout_rate,
        }


class InventoryService:
    """Lee inventarios Excel y calcula ISI/TQS sin alterar los originales."""

    REQUIRED_COLUMNS = {
        "date": "Fecha",
        "product": "Producto",
        "category": "Categoría",
        "sold_quantity": "Cantidad_Vendida",
        "stock_initial": "Stock_Inicial",
        "entries": "Entradas",
        "stock_available": "Stock_Disponible",
        "stock_final": "Stock_Final",
        "stockout_days": "Dias_Sin_Stock",
    }
    NUMERIC_COLUMNS = (
        "sold_quantity",
        "stock_initial",
        "entries",
        "stock_available",
        "stock_final",
        "stockout_days",
    )

    def __init__(self, storage_directory: str | Path | None = None):
        self.store = InventoryWorkbookStore(storage_directory) if storage_directory is not None else None

    def import_inventory(self, file_path: str | Path) -> InventoryImportResult:
        """Valida un Excel y, cuando hay almacén, guarda una copia inmutable."""
        source = self._validate_source_path(file_path)
        dataframe = self._read_inventory(source)
        stored: StoredInventoryFile | None = self.store.persist(source) if self.store is not None else None
        output_path = stored.path if stored is not None else source
        return InventoryImportResult(
            inventory_id=stored.inventory_id if stored is not None else None,
            stored_path=output_path,
            original_name=source.name,
            records=len(dataframe),
            categories=int(dataframe["category"].nunique()),
            start_date=dataframe["date"].min().date(),
            end_date=dataframe["date"].max().date(),
            reconciliation_rows=self._reconciliation_row_count(dataframe),
        )

    def calculate_category_metrics(
        self,
        file_path: str | Path,
        start_date: date | datetime | str | None = None,
        end_date: date | datetime | str | None = None,
    ) -> list[InventoryCategoryMetrics]:
        """Calcula SI, EN, SF, CD, ISI, DQS, DD y TQS por categoría.

        El rango se aplica de forma inclusiva. Si no se especifica, se evalúa
        todo el periodo contenido en el archivo importado.
        """
        dataframe = self._read_inventory(self._validate_source_path(file_path))
        start = self._parse_bound(start_date, "inicio")
        end = self._parse_bound(end_date, "fin")
        if start is not None and end is not None and end < start:
            raise ValueError("La fecha final del periodo no puede ser anterior a la fecha inicial.")

        if start is not None:
            dataframe = dataframe[dataframe["date"] >= pd.Timestamp(start)]
        if end is not None:
            dataframe = dataframe[dataframe["date"] <= pd.Timestamp(end)]
        if dataframe.empty:
            raise ValueError("No hay registros de inventario en el periodo seleccionado.")

        period_start = start or dataframe["date"].min().date()
        period_end = end or dataframe["date"].max().date()
        calendar_days = (period_end - period_start).days + 1

        metrics: list[InventoryCategoryMetrics] = []
        for category, rows in dataframe.groupby("category", sort=True):
            rows = rows.sort_values(["product", "date"], kind="stable").copy()
            # Cada producto puede tener muchas filas diarias. Para un indicador
            # de período, el saldo inicial y final se toman una sola vez: la
            # primera y última observación disponibles de dicho producto.
            # Las entradas sí son movimientos, por lo que se suman.
            product_groups = list(rows.groupby("product", sort=False))
            stock_initial = float(sum(group.iloc[0]["stock_initial"] for _, group in product_groups))
            entries = float(sum(group["entries"].sum() for _, group in product_groups))
            stock_final = float(sum(group.iloc[-1]["stock_final"] for _, group in product_groups))
            demand_quantity = stock_initial + entries - stock_final
            available = stock_initial + entries
            inventory_output_index = (demand_quantity / available * 100) if available else 0.0
            registered_days = int(rows["date"].nunique())
            evaluated_days = calendar_days
            coverage_rate = (registered_days / evaluated_days * 100) if evaluated_days else 0.0
            stockout_days = int(rows.loc[rows["stockout_days"] > 0, "date"].nunique())
            stockout_rate = (stockout_days / evaluated_days * 100) if evaluated_days else 0.0

            metrics.append(
                InventoryCategoryMetrics(
                    category=str(category),
                    start_date=rows["date"].min().date(),
                    end_date=rows["date"].max().date(),
                    records=len(rows),
                    stock_initial=stock_initial,
                    entries=entries,
                    stock_final=stock_final,
                    demand_quantity=demand_quantity,
                    inventory_output_index=inventory_output_index,
                    stockout_days=stockout_days,
                    evaluated_days=evaluated_days,
                    registered_days=registered_days,
                    coverage_rate=coverage_rate,
                    stockout_rate=stockout_rate,
                )
            )
        return metrics

    def reconciliation_summary(
        self,
        file_path: str | Path,
        start_date: date | datetime | str | None = None,
        end_date: date | datetime | str | None = None,
    ) -> dict[str, float | int | bool]:
        """Resume diferencias entre ventas registradas y balance de stock.

        La diferencia no invalida automáticamente un kardex: puede obedecer a
        mermas, devoluciones o ajustes aún no registrados. Se conserva como
        advertencia para que el usuario no presente el indicador como cerrado
        sin revisar la fuente de la tienda.
        """
        dataframe = self._read_inventory(self._validate_source_path(file_path))
        start = self._parse_bound(start_date, "inicio")
        end = self._parse_bound(end_date, "fin")
        if start is not None and end is not None and end < start:
            raise ValueError("La fecha final del periodo no puede ser anterior a la fecha inicial.")
        if start is not None:
            dataframe = dataframe[dataframe["date"] >= pd.Timestamp(start)]
        if end is not None:
            dataframe = dataframe[dataframe["date"] <= pd.Timestamp(end)]
        if dataframe.empty:
            raise ValueError("No hay registros de inventario en el periodo seleccionado.")

        product_groups = list(
            dataframe.sort_values(["category", "product", "date"], kind="stable").groupby(
                ["category", "product"], sort=False
            )
        )
        balance_demand = float(
            sum(
                group.iloc[0]["stock_initial"] + group["entries"].sum() - group.iloc[-1]["stock_final"]
                for _, group in product_groups
            )
        )
        sold_total = float(dataframe["sold_quantity"].sum())
        difference = sold_total - balance_demand
        row_count = self._reconciliation_row_count(dataframe)
        return {
            "rows_with_difference": row_count,
            "sales_total": sold_total,
            "balance_demand": balance_demand,
            "difference": difference,
            "reconciled": row_count == 0 and bool(np.isclose(difference, 0.0, rtol=0, atol=1e-6)),
        }

    @classmethod
    def _read_inventory(cls, file_path: Path) -> pd.DataFrame:
        try:
            raw = pd.read_excel(file_path, sheet_name=0, dtype=object)
        except (OSError, ValueError, ImportError, BadZipFile) as exc:
            raise ValueError(f"No se pudo leer el archivo de inventario: {exc}") from exc

        columns: dict[str, object] = {}
        for source_column in raw.columns:
            canonical = cls._canonical_column(source_column)
            if canonical is None:
                continue
            if canonical in columns:
                raise ValueError(f"El inventario repite la columna requerida '{cls.REQUIRED_COLUMNS[canonical]}'.")
            columns[canonical] = source_column

        missing = [label for key, label in cls.REQUIRED_COLUMNS.items() if key not in columns]
        if missing:
            raise ValueError("El inventario debe incluir las columnas: " + ", ".join(missing) + ".")

        dataframe = pd.DataFrame({key: raw[source] for key, source in columns.items()})
        dataframe = dataframe.dropna(how="all").copy()
        if dataframe.empty:
            raise ValueError("El inventario no contiene registros para procesar.")

        original_rows = dataframe.index
        parsed_dates = pd.to_datetime(dataframe["date"], errors="coerce", dayfirst=True)
        if parsed_dates.isna().any():
            raise ValueError(cls._invalid_row_message("Fecha", original_rows[parsed_dates.isna()]))
        dataframe["date"] = parsed_dates.dt.normalize()

        for column, label in (("product", "Producto"), ("category", "Categoría")):
            raw_text = dataframe[column]
            text = raw_text.fillna("").astype(str).str.strip()
            invalid = raw_text.isna() | text.eq("")
            if invalid.any():
                raise ValueError(cls._invalid_row_message(label, original_rows[invalid]))
            dataframe[column] = text

        for column in cls.NUMERIC_COLUMNS:
            numeric = cls._to_numeric(dataframe[column])
            invalid = numeric.isna()
            if invalid.any():
                raise ValueError(cls._invalid_row_message(cls.REQUIRED_COLUMNS[column], original_rows[invalid]))
            if (numeric < 0).any():
                rows = cls._excel_row_numbers(original_rows[numeric < 0])
                raise ValueError(
                    f"La columna '{cls.REQUIRED_COLUMNS[column]}' no puede tener valores negativos "
                    f"en las filas {rows}."
                )
            dataframe[column] = numeric.astype(float)

        # Cada fila representa el movimiento de un producto en un día. Antes
        # de sumar por categoría se comprueba que el balance básico de stock
        # sea coherente: no puede existir más stock final que disponible ni un
        # stock disponible distinto de SI + EN. La cantidad vendida puede no
        # coincidir exactamente por ajustes, mermas o devoluciones; por eso no
        # se rechaza ese caso sin una columna explícita de ajustes.
        expected_available = dataframe["stock_initial"] + dataframe["entries"]
        available_mismatch = ~np.isclose(
            dataframe["stock_available"], expected_available, rtol=0, atol=1e-6
        )
        if available_mismatch.any():
            rows = cls._excel_row_numbers(original_rows[available_mismatch])
            raise ValueError(
                "La columna 'Stock_Disponible' debe ser igual a Stock_Inicial + Entradas "
                f"en las filas {rows}."
            )
        final_above_available = dataframe["stock_final"] > dataframe["stock_available"] + 1e-6
        if final_above_available.any():
            rows = cls._excel_row_numbers(original_rows[final_above_available])
            raise ValueError(
                "La columna 'Stock_Final' no puede ser mayor que Stock_Disponible "
                f"en las filas {rows}."
            )

        return dataframe.reset_index(drop=True)

    @staticmethod
    def _reconciliation_row_count(dataframe: pd.DataFrame) -> int:
        """Cuenta filas donde venta y saldo no cierran sin bloquear la carga."""
        expected_final = dataframe["stock_available"] - dataframe["sold_quantity"]
        mismatch = ~np.isclose(dataframe["stock_final"], expected_final, rtol=0, atol=1e-6)
        return int(mismatch.sum())

    @staticmethod
    def _validate_source_path(file_path: str | Path) -> Path:
        path = Path(file_path)
        if path.suffix.lower() != ".xlsx":
            raise ValueError("El inventario debe estar en formato Excel (.xlsx).")
        if not path.exists() or not path.is_file():
            raise FileNotFoundError("No se encontró el archivo de inventario seleccionado.")
        if path.stat().st_size == 0:
            raise ValueError("El archivo de inventario está vacío.")
        return path

    @classmethod
    def _canonical_column(cls, value: object) -> str | None:
        # Excel puede presentar una cabecera con la tilde dañada como
        # ``Categor�a``. Sustituir el carácter de reemplazo antes de quitar
        # tildes conserva la compatibilidad con el archivo disponible.
        text = str(value).strip().replace("�", "i")
        text = unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode().lower()
        key = re.sub(r"[^a-z0-9]+", "", text)
        aliases = {
            "fecha": "date",
            "producto": "product",
            "categoria": "category",
            "cantidadvendida": "sold_quantity",
            "stockinicial": "stock_initial",
            "entradas": "entries",
            "stockdisponible": "stock_available",
            "stockfinal": "stock_final",
            "diassinstock": "stockout_days",
        }
        return aliases.get(key)

    @staticmethod
    def _to_numeric(values: pd.Series) -> pd.Series:
        def normalise(value: object) -> object:
            if not isinstance(value, str):
                return value
            text = value.strip().replace(" ", "")
            if text.count(",") == 1 and text.count(".") == 0:
                return text.replace(",", ".")
            if text.count(",") and text.count("."):
                if text.rfind(",") > text.rfind("."):
                    return text.replace(".", "").replace(",", ".")
                return text.replace(",", "")
            return text

        return pd.to_numeric(values.map(normalise), errors="coerce")

    @staticmethod
    def _parse_bound(value: date | datetime | str | None, label: str) -> date | None:
        if value is None or value == "":
            return None
        try:
            if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
                return datetime.strptime(value.strip(), "%Y-%m-%d").date()
            return pd.to_datetime(value, errors="raise", dayfirst=True).date()
        except (TypeError, ValueError) as exc:
            raise ValueError(f"La fecha de {label} del periodo no es válida.") from exc

    @staticmethod
    def _invalid_row_message(column: str, indexes: pd.Index | list[object]) -> str:
        return f"La columna '{column}' tiene un dato inválido en la fila {InventoryService._excel_row_numbers(indexes)}."

    @staticmethod
    def _excel_row_numbers(indexes: pd.Index | list[object]) -> str:
        excel_rows = [str(int(index) + 2) for index in list(indexes)[:5]]
        suffix = ", ..." if len(indexes) > 5 else ""
        return ", ".join(excel_rows) + suffix
