"""Reconstruye temporalmente el historial desde los 168 reportes POS.

No modifica los archivos activos de la aplicación. Comprueba que el CSV y el
archivo de mezcla por producto coinciden con las fuentes Excel.
"""

from __future__ import annotations

import tempfile
from pathlib import Path
import sys

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from application.services.catalog_service import CatalogService
from application.services.category_history_rebuild_service import CategoryHistoryRebuildService
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository
from infrastructure.repositories.csv_repository import CSVDemandRepository


def main():
    reports = sorted((ROOT / "data").glob("reporte_ventas_*.xlsx"))
    assert len(reports) == 168, f"se esperaban 168 reportes y se encontraron {len(reports)}"
    active_history = pd.read_csv(ROOT / "data" / "historial_demanda.csv").sort_values(["date", "category"]).reset_index(drop=True)
    active_mix = pd.read_csv(ROOT / "data" / "productos_por_categoria.csv").sort_values(["category", "product"]).reset_index(drop=True)

    with tempfile.TemporaryDirectory(prefix="taully-source-audit-") as temporary_directory:
        folder = Path(temporary_directory)
        rebuilt_history = folder / "historial_demanda.csv"
        rebuilt_mix = folder / "productos_por_categoria.csv"
        service = CategoryHistoryRebuildService(
            demand_repo=CSVDemandRepository(rebuilt_history),
            catalog_service=CatalogService(ExcelCatalogRepository(ROOT / "data" / "catalogo_maestro.xlsx")),
            sale_reader=ExcelReader(),
            backup_directory=folder / "backups",
            uncategorized_sales_path=folder / "ventas_sin_categoria.csv",
            category_product_mix_path=rebuilt_mix,
        )
        result = service.rebuild(reports)
        rebuilt = pd.read_csv(rebuilt_history).sort_values(["date", "category"]).reset_index(drop=True)
        mix = pd.read_csv(rebuilt_mix).sort_values(["category", "product"]).reset_index(drop=True)
        pd.testing.assert_frame_equal(active_history, rebuilt, check_dtype=False, check_exact=True)
        pd.testing.assert_frame_equal(active_mix, mix, check_dtype=False, check_exact=True)

    print(
        "FUENTES VERIFICADAS: "
        f"{result.reports} reportes, {result.dates} días, {result.categories} categorías, "
        f"{result.demand_records} registros, {result.categorized_units:.0f} unidades, "
        f"sin productos no categorizados."
    )


if __name__ == "__main__":
    main()
