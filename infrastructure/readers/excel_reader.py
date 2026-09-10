import pandas as pd
import re
from typing import List
from datetime import datetime

from domain.interfaces.repositories import SaleReader
from domain.entities.sale import Sale
from infrastructure.utils.date_cleaner import clean_sales_dataframe

class ExcelReader(SaleReader):
    """
    Lee archivos Excel que siguen el formato del sistema POS:
    - Encabezado con 'FECHAI: dd/mm/yyyy' y 'FECHAF: dd/mm/yyyy'.
    - Tabla con columnas: PROD, DESC, CANT, TOTAL.
    """
    def read_sales(self, file_path: str) -> List[Sale]:
        # 1. Leer el archivo una primera vez para encontrar la fila de la tabla
        df_raw = pd.read_excel(file_path, header=None, dtype=str)

        fecha = None
        inicio_tabla = None

        for idx, row in df_raw.iterrows():
            fila_texto = ' '.join([str(x) for x in row.values])
            if 'FECHAI:' in fila_texto and fecha is None:
                match = re.search(r'FECHAI:\s*(\d{2}/\d{2}/\d{4})', fila_texto)
                if match:
                    fecha = datetime.strptime(match.group(1), '%d/%m/%Y')

            if 'PROD' in fila_texto and 'CANT' in fila_texto:
                inicio_tabla = idx + 1
                break

        if fecha is None:
            raise ValueError("No se encontró 'FECHAI' en el archivo Excel")
        if inicio_tabla is None:
            raise ValueError("No se encontró la tabla con 'PROD' y 'CANT'")

        # 2. Leer nuevamente el archivo saltando las filas hasta el inicio de la tabla
        #    usamos skiprows para omitir las filas de encabezado y solo leer la tabla
        #    Además, forzamos dtype=str y header=None
        df_ventas = pd.read_excel(
            file_path,
            skiprows=inicio_tabla,   # salta todas las filas antes de la tabla
            header=None,
            dtype=str,
            usecols="A:D"            # solo las primeras 4 columnas
        )
        # Asignamos nombres de columnas
        df_ventas.columns = ['PROD', 'DESC', 'CANT', 'TOTAL']

        # 3. Limpiar datos
        df_ventas = clean_sales_dataframe(df_ventas)

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
