from abc import ABC, abstractmethod
from typing import List, Optional
from domain.entities.sale import Sale
from domain.entities.demand import Demand
from domain.entities.product import Product

class SaleReader(ABC):
    """Contrato para lectores de archivos (Excel/PDF) que extraen ventas."""
    @abstractmethod
    def read_sales(self, file_path: str) -> List[Sale]:
        pass

class DemandRepository(ABC):
    """Contrato para almacenar y recuperar la demanda histórica."""
    @abstractmethod
    def save_demands(self, demands: List[Demand]) -> dict[str, int]:
        """Guarda demandas y devuelve el detalle de filas nuevas o actualizadas."""
        pass

    @abstractmethod
    def get_all_demands(self) -> List[Demand]:
        pass

    @abstractmethod
    def get_demands_by_date_range(self, start_date, end_date) -> List[Demand]:
        pass

class ProductCatalogRepository(ABC):
    """Contrato para acceder al catálogo de productos."""
    @abstractmethod
    def get_product(self, product_name: str) -> Optional[Product]:
        pass

    @abstractmethod
    def get_all_products(self) -> List[Product]:
        pass
