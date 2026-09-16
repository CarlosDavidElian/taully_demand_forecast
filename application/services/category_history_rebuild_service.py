"""Reconstrucción trazable del historial diario por categoría comercial."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import csv
from pathlib import Path
import shutil
from typing import Iterable

from application.services.catalog_service import CatalogService
from domain.entities.demand import Demand
from domain.interfaces.repositories import SaleReader
from infrastructure.repositories.csv_repository import CSVDemandRepository


@dataclass(frozen=True)
class CategoryHistoryRebuildResult:
    """Resumen verificable de una reconstrucción del historial."""

    reports: int
    dates: int
    categories: int
    demand_records: int
    source_units: float
    categorized_units: float
    uncategorized_units: float
    uncategorized_products: int
    backup_path: Path | None
    uncategorized_sales_path: Path
    category_product_mix_path: Path

    @property
    def categorized_unit_share(self) -> float:
        if self.source_units == 0:
            return 0.0
        return self.categorized_units / self.source_units


class CategoryHistoryRebuildService:
    """Genera una serie diaria completa a partir de reportes POS.

    Cada fecha queda representada para cada categoría activa del catálogo. Una
    venta ausente se convierte en cero solo porque el reporte de ese día existe;
    así el modelo distingue entre ausencia de venta y ausencia de información.
    """

    def __init__(
        self,
        demand_repo: CSVDemandRepository,
        catalog_service: CatalogService,
        sale_reader: SaleReader,
        backup_directory: Path,
        uncategorized_sales_path: Path,
        category_product_mix_path: Path,
    ):
        self.demand_repo = demand_repo
        self.catalog_service = catalog_service
        self.sale_reader = sale_reader
        self.backup_directory = backup_directory
        self.uncategorized_sales_path = uncategorized_sales_path
        self.category_product_mix_path = category_product_mix_path

    def rebuild(self, report_paths: Iterable[Path]) -> CategoryHistoryRebuildResult:
        paths = sorted(Path(path) for path in report_paths)
        if not paths:
            raise ValueError("No se encontraron reportes POS para reconstruir el historial.")

        categories = sorted({product.category for product in self.catalog_service.get_all_products()})
        if not categories:
            raise ValueError("El catálogo activo no contiene categorías.")

        category_totals: defaultdict[tuple[datetime, str], float] = defaultdict(float)
        category_product_totals: defaultdict[tuple[str, str], float] = defaultdict(float)
        uncategorized_totals: defaultdict[tuple[datetime, str], float] = defaultdict(float)
        source_dates: set[datetime] = set()
        source_units = 0.0
        categorized_units = 0.0

        for path in paths:
            sales = self.sale_reader.read_sales(str(path))
            if not sales:
                raise ValueError(f"El reporte '{path.name}' no contiene ventas válidas.")

            report_dates = {sale.date.replace(hour=0, minute=0, second=0, microsecond=0) for sale in sales}
            if len(report_dates) != 1:
                raise ValueError(f"El reporte '{path.name}' contiene más de una fecha de venta.")
            report_date = report_dates.pop()
            if report_date in source_dates:
                raise ValueError(f"Hay más de un reporte para la fecha {report_date:%Y-%m-%d}.")
            source_dates.add(report_date)

            for sale in sales:
                source_units += float(sale.quantity)
                product = self.catalog_service.get_product(sale.product_name)
                if product is None:
                    uncategorized_totals[(report_date, sale.product_name)] += float(sale.quantity)
                    continue
                category_totals[(report_date, product.category)] += float(sale.quantity)
                category_product_totals[(product.category, product.product_name)] += float(sale.quantity)
                categorized_units += float(sale.quantity)

        ordered_dates = sorted(source_dates)
        self._validate_contiguous_dates(ordered_dates)
        demands = [
            Demand(date, category, category_totals[(date, category)])
            for date in ordered_dates
            for category in categories
        ]

        # Se completan todos los archivos de control antes de reemplazar el
        # historial actual. La copia permite recuperar el estado previo.
        backup_path = self._backup_existing_history()
        self._write_uncategorized_sales(uncategorized_totals)
        self._write_category_product_mix(category_product_totals)
        self.demand_repo.replace_demands(demands)

        uncategorized_units = source_units - categorized_units
        return CategoryHistoryRebuildResult(
            reports=len(paths),
            dates=len(ordered_dates),
            categories=len(categories),
            demand_records=len(demands),
            source_units=source_units,
            categorized_units=categorized_units,
            uncategorized_units=uncategorized_units,
            uncategorized_products=len({product for _, product in uncategorized_totals}),
            backup_path=backup_path,
            uncategorized_sales_path=self.uncategorized_sales_path,
            category_product_mix_path=self.category_product_mix_path,
        )

    @staticmethod
    def _validate_contiguous_dates(dates: list[datetime]) -> None:
        if not dates:
            raise ValueError("No se encontraron fechas de venta en los reportes.")
        expected_days = (dates[-1].date() - dates[0].date()).days + 1
        if len(dates) != expected_days:
            missing_dates = []
            current = dates[0]
            known_dates = set(dates)
            while current <= dates[-1] and len(missing_dates) < 5:
                if current not in known_dates:
                    missing_dates.append(current.strftime("%Y-%m-%d"))
                current += timedelta(days=1)
            suffix = ", ".join(missing_dates)
            raise ValueError(f"Faltan reportes diarios para reconstruir el historial: {suffix}.")

    def _backup_existing_history(self) -> Path | None:
        current_history = self.demand_repo.file_path
        if not current_history.exists():
            return None
        self.backup_directory.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        backup_path = self.backup_directory / f"{current_history.stem}_antes_reconstruccion_{timestamp}.csv"
        shutil.copy2(current_history, backup_path)
        return backup_path

    def _write_uncategorized_sales(self, totals: dict[tuple[datetime, str], float]) -> None:
        self.uncategorized_sales_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path = self.uncategorized_sales_path.with_suffix(".tmp")
        try:
            with temporary_path.open("w", newline="", encoding="utf-8") as output:
                writer = csv.DictWriter(output, fieldnames=["date", "product", "quantity"])
                writer.writeheader()
                for (sale_date, product), quantity in sorted(totals.items()):
                    writer.writerow(
                        {
                            "date": sale_date.strftime("%Y-%m-%d"),
                            "product": product,
                            "quantity": quantity,
                        }
                    )
            temporary_path.replace(self.uncategorized_sales_path)
        finally:
            temporary_path.unlink(missing_ok=True)

    def _write_category_product_mix(self, totals: dict[tuple[str, str], float]) -> None:
        self.category_product_mix_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path = self.category_product_mix_path.with_suffix(".tmp")
        try:
            with temporary_path.open("w", newline="", encoding="utf-8") as output:
                writer = csv.DictWriter(output, fieldnames=["category", "product", "historical_quantity"])
                writer.writeheader()
                for (category, product), quantity in sorted(
                    totals.items(), key=lambda item: (item[0][0], -item[1], item[0][1])
                ):
                    writer.writerow(
                        {
                            "category": category,
                            "product": product,
                            "historical_quantity": quantity,
                        }
                    )
            temporary_path.replace(self.category_product_mix_path)
        finally:
            temporary_path.unlink(missing_ok=True)
