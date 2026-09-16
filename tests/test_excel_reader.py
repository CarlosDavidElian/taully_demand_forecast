from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from infrastructure.readers.excel_reader import ExcelReader


class ExcelReaderTests(unittest.TestCase):
    def test_reads_daily_sales_report(self):
        with TemporaryDirectory() as folder:
            report = Path(folder) / "reporte_ventas_2026-04-01.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "REPORTE"
            sheet.append(["REPORTE DE VENTAS - 01/04/2026"])
            sheet.append([])
            sheet.append(["VENTAS DEL DÍA", "TOTAL"])
            sheet.append(["Total", 10])
            sheet.append([])
            sheet.append(["PRODUCTO", "CANTIDAD VENDIDA", "CATEGORIA", "P. UNITARIO", "TOTAL"])
            sheet.append(["Producto A", 3, "ABARROTES", 2.5, 7.5])
            sheet.append(["Producto B", 2, "BEBIDAS", 4, 8])
            workbook.save(report)

            sales = ExcelReader().read_sales(str(report))

        self.assertEqual(len(sales), 2)
        self.assertEqual(sales[0].date, datetime(2026, 4, 1))
        self.assertEqual(sales[0].product_name, "Producto A")
        self.assertEqual(sales[0].quantity, 3)
        self.assertEqual(sales[0].total, 7.5)

    def test_keeps_compatibility_with_pos_report(self):
        with TemporaryDirectory() as folder:
            report = Path(folder) / "pos.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["FECHAI: 13/09/2026"])
            sheet.append([])
            sheet.append(["PROD", "DESC", "CANT", "TOTAL"])
            sheet.append(["A01", "Producto POS", 4, 12])
            workbook.save(report)

            sales = ExcelReader().read_sales(str(report))

        self.assertEqual(len(sales), 1)
        self.assertEqual(sales[0].date, datetime(2026, 9, 13))
        self.assertEqual(sales[0].product_name, "A01")
