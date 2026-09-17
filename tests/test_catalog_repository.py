from __future__ import annotations

from pathlib import Path
import unittest

from infrastructure.repositories.catalog_repository import ExcelCatalogRepository


class ExcelCatalogRepositoryTests(unittest.TestCase):
    def test_reads_product_sheet_and_normalises_lookup_and_cost(self):
        catalog_path = Path(__file__).resolve().parents[1] / "data" / "catalogo_maestro.xlsx"
        repository = ExcelCatalogRepository(catalog_path)

        products = repository.get_all_products()
        self.assertGreater(len(products), 0)

        source_product = products[0]
        product = repository.get_product(f"  {source_product.product_name.lower()}  ")
        self.assertIsNotNone(product)
        self.assertEqual(product.product_name, source_product.product_name)
        self.assertEqual(product.cost, source_product.cost)

        # Los productos con costo numérico deben mantenerse disponibles para
        # la clasificación, incluso cuando el catálogo se actualice.
        self.assertGreaterEqual(product.cost, 0.0)
