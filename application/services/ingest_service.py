from typing import List
from pathlib import Path

from domain.entities.sale import Sale
from domain.entities.demand import Demand
from domain.interfaces.repositories import SaleReader, DemandRepository, ProductCatalogRepository
from infrastructure.readers.excel_reader import ExcelReader
from infrastructure.readers.pdf_reader import PDFReader

class IngestService:
    def __init__(
        self,
        catalog_repo: ProductCatalogRepository,
        demand_repo: DemandRepository
    ):
        self.catalog_repo = catalog_repo
        self.demand_repo = demand_repo

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

        # 3. Consolidar por producto. Cada presentación se mantiene separada
        #    para que el pronóstico represente, por ejemplo, "GLORIA AZUL 400"
        #    y "GLORIA UHT ROJA" como productos diferentes.
        demands_dict = {}
        for sale in sales:
            product_name = sale.product_name

            # Agrupar por (fecha, producto) sumando cantidades cuando el mismo
            # producto aparece más de una vez en un reporte.
            key = (sale.date, product_name)
            if key in demands_dict:
                demands_dict[key] += sale.quantity
            else:
                demands_dict[key] = sale.quantity

        # 4. Convertir a entidades Demand
        demands = [
            Demand(date=date, category=cat, quantity=qty)
            for (date, cat), qty in demands_dict.items()
        ]

        # 5. Guardar en el repositorio (historial)
        self.demand_repo.save_demands(demands)

        return demands
