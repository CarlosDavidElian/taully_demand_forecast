from typing import Any, List, Optional

from domain.entities.product import Product
from domain.interfaces.repositories import ProductCatalogRepository

class CatalogService:
    def __init__(self, catalog_repo: ProductCatalogRepository):
        self.catalog_repo = catalog_repo

    def get_all_products(self) -> List[Product]:
        return self.catalog_repo.get_all_products()

    def get_product(self, name: str) -> Optional[Product]:
        return self.catalog_repo.get_product(name)

    def get_summary(self) -> dict[str, Any]:
        """Resume las clasificaciones disponibles en el catálogo activo."""
        products = self.get_all_products()
        return {
            "products": len(products),
            "families": len({product.family for product in products}),
            "categories": len({product.category for product in products}),
            "brands": len({product.brand for product in products if product.brand}),
        }

    def get_categories(self) -> list[str]:
        """Devuelve las categorías comerciales activas del catálogo."""
        return sorted({product.category for product in self.get_all_products()})
