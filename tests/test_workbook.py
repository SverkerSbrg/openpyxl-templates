from unittest import TestCase

from openpyxl_templates.templated_workbook import TemplatedWorkbook
from tests.test_table_sheet.test_table_sheet import TestTemplatedSheet


class TestTemplatedWorkbook(TemplatedWorkbook):
    sheet1 = TestTemplatedSheet(sheetname="Test")
    sheet2 = TestTemplatedSheet(sheetname="Test2")


class TemplatedWorkbookTests(TestCase):
    def setUp(self):
        self.wb = TestTemplatedWorkbook()

    def test_templated_sheets(self):
        self.assertEqual(
            [self.wb.sheet1, self.wb.sheet2],
            self.wb.templated_sheets
        )

    def test_templated_sheets_workbook(self):
        for sheet in self.wb.templated_sheets:
            self.assertEqual(sheet.workbook, self.wb)

    def test_exists_false(self):
        self.assertFalse(self.wb.sheet1.exists)

    def test_exists_true(self):
        # Automatically create on access.
        ws = self.wb.sheet1.worksheet
        self.assertTrue(self.wb.sheet1.exists)

    def test_sheet_index(self):
        self.wb.remove_all_sheets()

        ws = self.wb.sheet1.worksheet
        ws = self.wb.sheet2.worksheet

        self.assertEqual(0, self.wb.sheet1.sheet_index)
        self.assertEqual(1, self.wb.sheet2.sheet_index)
