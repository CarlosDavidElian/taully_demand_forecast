import pandas as pd
from typing import List, Optional
from pathlib import Path

from domain.interfaces.repositories import ProductCatalogRepository
from domain.entities.product import Product
from config.settings import CATALOG_FILE

class ExcelCatalogRepository(ProductCatalogRepository):
    def __init__(self):
        self._products = []
        self._load_catalog()

    def _load_catalog(self):
        # 1. Verificar que el archivo existe
        if not CATALOG_FILE.exists():
            raise FileNotFoundError(f"Catálogo no encontrado en {CATALOG_FILE}")
        
        print(f"INFO Leyendo catálogo desde: {CATALOG_FILE}")
        print(f"INFO Tamaño del archivo: {CATALOG_FILE.stat().st_size} bytes")
        
        # 2. Intentar leer con openpyxl
        try:
            df = pd.read_excel(CATALOG_FILE, header=0, dtype=str, engine='openpyxl')
        except Exception as e:
            print(f"ERROR al leer con openpyxl: {e}")
            # Fallback: leer sin especificar motor
            df = pd.read_excel(CATALOG_FILE, header=0, dtype=str)
        
        print(f"OK Columnas encontradas: {df.columns.tolist()}")
        print(f"OK Filas encontradas: {len(df)}")
        if len(df) > 0:
            print("Primeras 3 filas:\n", df.head(3))
        
        # 3. Normalizar nombres de columnas
        df.columns = df.columns.str.upper().str.strip()
        
        required_cols = ['PRODUCTO', 'FAMILIA', 'CATEGORIA']
        for col in required_cols:
            if col not in df.columns:
                raise ValueError(f"El catálogo debe tener la columna {col}")
        
        # 4. Cargar productos
        self._products = []
        for _, row in df.iterrows():
            product_name = str(row['PRODUCTO']).strip()
            family = str(row['FAMILIA']).strip()
            category = str(row['CATEGORIA']).strip()
            brand = str(row.get('MARCA', '')).strip()
            cost_str = str(row.get('COSTO', '0')).replace(',', '.').strip()
            try:
                cost = float(cost_str) if cost_str else 0.0
            except ValueError:
                cost = 0.0
            
            product = Product(
                product_name=product_name,
                family=family,
                category=category,
                brand=brand,
                cost=cost
            )
            self._products.append(product)
        
        print(f"OK {len(self._products)} productos cargados correctamente")

    def get_product(self, product_name: str) -> Optional[Product]:
        for p in self._products:
            if p.product_name == product_name:
                return p
        return None

    def get_all_products(self) -> List[Product]:
        return self._products
