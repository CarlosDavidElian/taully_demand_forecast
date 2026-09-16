from __future__ import annotations

from pathlib import Path
import unittest

from infrastructure.readers.excel_reader import ExcelReader


class ExcelReaderTests(unittest.TestCase):
    def test_reads_known_pos_report(self):
        report = Path(__file__).resolve().parents[1] / "data" / "Reporte_Taully_2026-09-13.xlsx"
        sales = ExcelReader().read_sales(str(report))

        self.assertEqual(len(sales), 14)
        self.assertEqual(sales[0].date.strftime("%Y-%m-%d"), "2026-09-13")
        self.assertTrue(all(sale.quantity > 0 for sale in sales))
