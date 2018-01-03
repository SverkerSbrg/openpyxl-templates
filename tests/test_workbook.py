from unittest import TestCase

from openpyxl_templates.table_sheet import TableSheet, TableColumn
from openpyxl_templates.templated_workbook import TemplatedWorkbook, SheetnamesNotUnique, MultipleActiveSheets


class TestTemplatedSheet(TableSheet):
    column1 = TableColumn(header="column1")
    column2 = TableColumn(header="column2")
    column3 = TableColumn(header="column3")


class TestTemplatedWorkbook(TemplatedWorkbook):
    sheet1 = TestTemplatedSheet(sheetname="Custom sheetname")
    sheet2 = TestTemplatedSheet()



class TemplatedWorkbookTests(TestCase):
    def setUp(self):
        self.wb = TestTemplatedWorkbook()

    def test_templated_sheets(self):
        self.assertEqual(
            [self.wb.sheet1, self.wb.sheet2],
            self.wb.templated_sheets
        )

    def test_dynamic_templated_sheets(self):
        sheet3 = TestTemplatedSheet(sheetname="Dynamic sheet")

        wb = TestTemplatedWorkbook(templated_sheets=[sheet3,])

        self.assertEqual(
            [wb.sheet1, wb.sheet2, sheet3],
            list(wb.templated_sheets)
        )


    def test_templated_sheets_workbook(self):
        for sheet in self.wb.templated_sheets:
            self.assertEqual(sheet.workbook, self.wb.workbook)

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

    def test_sheetname_from_workbook_attribute(self):
        self.assertEqual(self.wb.sheet1.sheetname, "Custom sheetname")
        self.assertEqual(self.wb.sheet2.sheetname, "sheet2")

    def test_sheetnames_not_unique(self):
        class SheetnamesNotUniqueWorkbook(TemplatedWorkbook):
            sheet1 = TestTemplatedSheet(sheetname="Test")
            sheet2 = TestTemplatedSheet(sheetname="Test")

        with self.assertRaises(SheetnamesNotUnique):
            SheetnamesNotUniqueWorkbook()

    def test_sort(self):
        class TestTemplatedWorkbook(TemplatedWorkbook):
            sheet1 = TestTemplatedSheet()
            sheet2 = TestTemplatedSheet()

        wb = TestTemplatedWorkbook()

        wb.remove_all_sheets()

        wb.create_sheet("ordinary_sheet1")
        wb.create_sheet("ordinary_sheet2")

        ws = wb.sheet2.worksheet
        ws = wb.sheet1.worksheet

        self.assertEqual(
            tuple(wb.sheetnames),
            ("ordinary_sheet1", "ordinary_sheet2", "sheet2", "sheet1")
        )
        wb.sort_worksheets()
        self.assertEqual(
            tuple(wb.sheetnames),
            ("sheet1", "sheet2", "ordinary_sheet1", "ordinary_sheet2")
        )

    def test_multiple_active(self):
        class MultipleActiveWorkbook(TemplatedWorkbook):
            sheet1 = TestTemplatedSheet(active=True)
            sheet2 = TestTemplatedSheet(active=True)

        with self.assertRaises(MultipleActiveSheets):
            MultipleActiveWorkbook()

    # def test_asdf(self):
    #     self.wb.create_sheet("asdf")
    #     ws = self.wb["asdf"]
    #     x = self.wb["asdf"].__iter__()