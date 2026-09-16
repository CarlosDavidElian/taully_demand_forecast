import pandas as pd
import re
from typing import List
from datetime import datetime
import unicodedata

from domain.interfaces.repositories import SaleReader
from domain.entities.sale import Sale
from infrastructure.utils.date_cleaner import clean_sales_dataframe

class ExcelReader(SaleReader):
    """
    Lee los dos formatos de reportes Excel que usa el proyecto:
    - POS: encabezado ``FECHAI`` y tabla ``PROD`` / ``CANT``.
    - Reporte diario: título ``REPORTE DE VENTAS - dd/mm/aaaa`` y tabla
      ``PRODUCTO`` / ``CANTIDAD VENDIDA``.
    """
    def read_sales(self, file_path: str) -> List[Sale]:
        # Leemos la hoja completa una sola vez. El reporte POS suele tener
        # información antes de la tabla y no siempre fija sus columnas en A:D.
        df_raw = pd.read_excel(file_path, header=None, dtype=object)
        fecha = self._extract_start_date(df_raw)
        header_index, column_indexes = self._find_table_header(df_raw)

        selected = pd.DataFrame(
            {
                "PROD": df_raw.iloc[header_index + 1 :, column_indexes["PROD"]],
                "CANT": df_raw.iloc[header_index + 1 :, column_indexes["CANT"]],
            }
        )
        if "TOTAL" in column_indexes:
            selected["TOTAL"] = df_raw.iloc[header_index + 1 :, column_indexes["TOTAL"]]
        else:
            # TOTAL no se usa para consolidar demanda; mantenerlo en cero
            # permite procesar reportes válidos que no lo muestran.
            selected["TOTAL"] = "0"

        # Limpiar datos
        df_ventas = selected
        df_ventas = clean_sales_dataframe(df_ventas)
        if df_ventas.empty:
            raise ValueError("El reporte Excel no contiene filas de venta válidas.")

        # Mapear a entidades Sale
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

    @staticmethod
    def _extract_start_date(dataframe: pd.DataFrame) -> datetime:
        """Encuentra la fecha en el encabezado POS o en el título diario."""
        patterns = (
            re.compile(r"FECHAI\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})", re.IGNORECASE),
            re.compile(r"REPORTE\s+DE\s+VENTAS\s*-\s*(\d{1,2}/\d{1,2}/\d{4})", re.IGNORECASE),
        )
        for _, row in dataframe.iterrows():
            values = [str(value).strip() for value in row.values if pd.notna(value)]
            row_text = " ".join(values)
            for pattern in patterns:
                match = pattern.search(row_text)
                if match:
                    return datetime.strptime(match.group(1), "%d/%m/%Y")

            # Algunos exportadores separan "FECHAI" y "12/09/2026" en
            # dos celdas. Solo aceptamos una fecha de la misma fila.
            if any(value.upper().rstrip(":") == "FECHAI" for value in values):
                for value in values:
                    date_match = re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", value)
                    if date_match:
                        return datetime.strptime(date_match.group(0), "%d/%m/%Y")
        raise ValueError("No se encontró una fecha de venta en el archivo Excel")

    @staticmethod
    def _find_table_header(dataframe: pd.DataFrame) -> tuple[int, dict[str, int]]:
        aliases = {
            "PROD": "PROD",
            "PRODUCTO": "PROD",
            "CANT": "CANT",
            "CANTIDAD VENDIDA": "CANT",
            "TOTAL": "TOTAL",
        }
        required = {"PROD", "CANT"}
        for row_index, row in dataframe.iterrows():
            headers: dict[str, int] = {}
            for column_index, value in enumerate(row.values):
                if pd.isna(value):
                    continue
                header = unicodedata.normalize("NFKD", str(value)).encode("ASCII", "ignore").decode().strip().upper()
                canonical_name = aliases.get(header)
                if canonical_name:
                    headers.setdefault(canonical_name, column_index)
            if required.issubset(headers):
                return row_index, headers
        raise ValueError("No se encontró una tabla con producto y cantidad vendida")
