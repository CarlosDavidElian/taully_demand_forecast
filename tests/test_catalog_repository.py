from __future__ import annotations

from pathlib import Path
import unittest

from infrastructure.repositories.catalog_repository import ExcelCatalogRepository


class ExcelCatalogRepositoryTests(unittest.TestCase):
    def test_reads_product_sheet_and_normalises_lookup_and_cost(self):
        catalog_path = Path(__file__).resolve().parents[1] / "data" / "catalogo_maestro.xlsx"
        repository = ExcelCatalogRepository(catalog_path)

        self.assertGreater(len(repository.get_all_products()), 0)

        product = repository.get_product("  gloria azul 400  ")
        self.assertIsNotNone(product)
        self.assertEqual(product.product_name, "GLORIA AZUL 400")

        # Este valor llega desde Excel como una fecha de 1900 si no se
        # convierte nuevamente al serial monetario.
        self.assertEqual(repository.get_product("AKIO AVENA 1.0").cost, 1.2)

        # Los productos con costo numérico deben mantenerse disponibles para
        # la clasificación, incluso cuando el catálogo se actualice.
        self.assertGreaterEqual(product.cost, 0.0)
