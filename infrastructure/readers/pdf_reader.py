import pandas as pd
import tabula
import pdfplumber
from typing import List

from domain.interfaces.repositories import SaleReader
from domain.entities.sale import Sale
from infrastructure.utils.date_extractor import extract_date_from_text
from infrastructure.utils.date_cleaner import clean_sales_dataframe

class PDFReader(SaleReader):
    """
    Lee archivos PDF que siguen el formato del sistema POS:
    - Encabezado con 'FECHAI: dd/mm/yyyy' y 'FECHAF: dd/mm/yyyy'.
    - Tabla con columnas: PROD, DESC, CANT, TOTAL.
    """
    def read_sales(self, file_path: str) -> List[Sale]:
        # 1. Extraer la fecha del PDF usando pdfplumber
        fecha = None
        with pdfplumber.open(file_path) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += (page.extract_text() or "") + "\n"
            fecha = extract_date_from_text(texto_completo)

        # 2. Extraer la tabla usando tabula (busca la primera tabla que tenga PROD y CANT)
        try:
            tablas = tabula.read_pdf(file_path, pages='all', multiple_tables=True, guess=False)
        except Exception as exc:
            raise ValueError("No se pudo leer la tabla de ventas del PDF.") from exc
        
        df_ventas = None
        for tabla in tablas:
            # Normalizar nombres de columnas
            tabla.columns = tabla.columns.str.upper().str.strip()
            if 'PROD' in tabla.columns and 'CANT' in tabla.columns:
                df_ventas = tabla[['PROD', 'CANT']].copy()
                df_ventas['TOTAL'] = tabla['TOTAL'] if 'TOTAL' in tabla.columns else '0'
                break
        
        if df_ventas is None:
            raise ValueError("No se encontró la tabla de ventas en el PDF")
        
        # 3. Limpiar los datos (quitar ANULADO, convertir a números)
        df_ventas = clean_sales_dataframe(df_ventas)
        if df_ventas.empty:
            raise ValueError("El PDF no contiene filas de venta válidas.")

        # 4. Mapear a entidades Sale
        sales = []
        for _, row in df_ventas.iterrows():
            sale = Sale(
                date=fecha,
                product_name=str(row['PROD']).strip(),
                quantity=float(row['CANT']),
                total=float(row['TOTAL'])
            )
            sales.append(sale)

        return sales
