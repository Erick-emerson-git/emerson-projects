"""Offline regression tests; no accounts or network required."""
from datetime import datetime
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook
from main import generate, EXTRA


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / 'report.xlsx'
        self.headers = ['ITEM', 'ORG', 'TOTAL INVENTORY', 'PO IN RECEIVING',
                        'TOTAL DEMAND CURRENT WEEK', 'TOTAL DEMAND WEEK 2',
                        'AWU', 'PRIORITY NUMBER']
        self.raw = [dict(zip(self.headers, ['0001', 'TFC', 10, 2, 8, 4, 2, '1.0']))]
        self.sql = [{'ORG': 'MMC', 'ITEM_NAME': '0001',
                     'NEWEST_TRANSACTION_DATE': datetime(2026, 10, 15, 13, 5),
                     'TOTAL_QTY_LAST_24H': -6}]

    def test_item_match_signed_qty_real_date_and_formulas(self):
        stats = generate(self.headers, self.raw, self.sql, self.path)
        self.assertEqual(stats['cross_org_matches'], 1)
        wb = load_workbook(self.path, data_only=True)
        ws = wb.active
        col = {c.value: c.column for c in ws[1]}
        self.assertEqual(ws.cell(2, col[EXTRA[1]]).value, -6)
        self.assertEqual(ws.cell(2, col[EXTRA[2]]).value, datetime(2026, 10, 15, 13, 5))
        self.assertEqual(ws.cell(2, col[EXTRA[0]]).value, 'REVIEW')  # strict >, not >=
        self.assertEqual(ws.cell(2, col[EXTRA[6]]).value, 8)
        self.assertIsNone(ws.cell(2, col[EXTRA[3]]).value)
        self.assertFalse(ws.sheet_view.showGridLines)
        self.assertEqual(ws.cell(2, col['PRIORITY NUMBER']).fill.fgColor.rgb, 'FFFF0000')
        self.assertEqual(ws.cell(2, col[EXTRA[-1]]).border.left.style, 'thin')
        wb.close()
        wb = load_workbook(self.path, data_only=False)
        self.assertIn('ABS(', wb.active.cell(2, col[EXTRA[6]]).value)
        self.assertEqual(len(wb._external_links), 0)
        wb.close()

    def test_org_match_does_not_use_other_organization(self):
        stats = generate(self.headers, self.raw, self.sql, self.path, 'org_item', 'na')
        self.assertEqual(stats['matched'], 0)
        wb = load_workbook(self.path, data_only=True)
        self.assertEqual(wb.active.cell(2, len(self.headers)+2).value, '#N/A')
        wb.close()

    def test_first_match_and_repeated_raw_rows_preserved(self):
        second = dict(self.sql[0], TOTAL_QTY_LAST_24H=-99)
        stats = generate(self.headers, self.raw*2, self.sql+[second], self.path)
        self.assertEqual(stats['duplicate_sql_keys'], 1)
        self.assertEqual(stats['matched'], 2)
        wb = load_workbook(self.path, data_only=True)
        self.assertEqual(wb.active.max_row, 3)
        self.assertEqual(wb.active.cell(3, len(self.headers)+2).value, -6)
        self.assertEqual(wb.active['A2'].value, '0001')
        wb.close()


if __name__ == '__main__':
    unittest.main()
