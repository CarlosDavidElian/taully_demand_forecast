from typing import List
from pathlib import Path

from application.services.catalog_service import CatalogService
from domain.entities.sale import Sale
from domain.entities.demand import Demand
from domain.interfaces.repositories import SaleReader, DemandRepository
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.readers.pdf_reader import PDFReader

class IngestService:
    def __init__(
        self,
        demand_repo: DemandRepository,
        catalog_service: CatalogService | None = None,
    ):
        self.demand_repo = demand_repo
        self.catalog_service = catalog_service
        self.last_save_summary = {"added": 0, "updated": 0, "unchanged": 0}
        self.last_catalog_summary = {"matched_products": 0, "unmatched_products": []}

    def process_file(self, file_path: str) -> List[Demand]:
        # 1. Seleccionar el reader adecuado según la extensión
        path = Path(file_path)
        if path.suffix.lower() in ['.xlsx', '.xls']:
            reader: SaleReader = ExcelReader()
        elif path.suffix.lower() == '.pdf':
            reader: SaleReader = PDFReader()
        else:
            raise ValueError("Formato no soportado. Use .xlsx, .xls o .pdf")

        # 2. Leer las ventas del archivo
        sales: List[Sale] = reader.read_sales(str(path))

        # 3. Consolidar por fecha + categoría comercial. La unidad de análisis
        #    del proyecto es la demanda diaria por categoría, no el SKU.
        category_totals: dict[tuple[object, str], float] = {}
        report_dates = set()
        matched_products: set[str] = set()
        unmatched_products: set[str] = set()
        matched_units = 0.0
        unmatched_units = 0.0
        for sale in sales:
            report_date = sale.date.replace(hour=0, minute=0, second=0, microsecond=0)
            report_dates.add(report_date)
            product_name = sale.product_name.strip()
            category = None
            if self.catalog_service is not None:
                catalog_product = self.catalog_service.get_product(product_name)
                if catalog_product is not None:
                    product_name = catalog_product.product_name
                    matched_products.add(product_name)
                    matched_units += float(sale.quantity)
                    category = catalog_product.category
                else:
                    unmatched_products.add(product_name)
                    unmatched_units += float(sale.quantity)

            if category is None:
                # Sin una categoría maestra no se puede atribuir la venta a un
                # grupo comercial. Se informa al usuario para completar el
                # catálogo, sin contaminar el modelo por categorías.
                continue

            key = (report_date, category)
            category_totals[key] = category_totals.get(key, 0.0) + float(sale.quantity)

        if self.catalog_service is None:
            raise ValueError("Se requiere un catálogo activo para consolidar la demanda por categoría.")
        categories = self.catalog_service.get_categories()
        demands = [
            Demand(date=report_date, category=category, quantity=category_totals.get((report_date, category), 0.0))
            for report_date in sorted(report_dates)
            for category in categories
        ]

        if not demands:
            raise ValueError("El reporte no contiene ventas válidas para incorporar.")

        # 5. Guardar en el repositorio (historial) y conservar el resultado
        #    para informar a la interfaz qué cambió realmente.
        self.last_save_summary = self.demand_repo.save_demands(demands)
        self.last_catalog_summary = {
            "matched_products": len(matched_products),
            "unmatched_products": sorted(unmatched_products),
            "matched_units": round(matched_units, 2),
            "unmatched_units": round(unmatched_units, 2),
        }

        return demands
