from unittest import TestCase

from openpyxl_templates.table_sheet.columns import TableColumn
from openpyxl_templates.table_sheet.table_sheet import TableSheet, ColumnHeadersNotUnique, NoTableColumns, \
    CannotHideOrGroupLastColumn, HeadersNotFound, MultipleFrozenColumns
from openpyxl_templates.templated_workbook import TemplatedWorkbook
from openpyxl_templates.utils import FakeCells


class TestTemplatedSheet(TableSheet):
    column1 = TableColumn(header="column1")
    column2 = TableColumn(header="column2")
    column3 = TableColumn(header="column3")


class TestTemplatedWorkbook(TemplatedWorkbook):
    sheet1 = TestTemplatedSheet()


class FakeTableSheet(TableSheet):
    column1 = TableColumn(header="column1")
    column2 = TableColumn(header="column2")
    column3 = TableColumn(header="column3")

    def __init__(self, *rows):
        self.fake_worksheet = (FakeCells(*row) for row in rows)

        super(FakeTableSheet, self).__init__(sheetname="fakesheet")

    @property
    def worksheet(self):
        return self.fake_worksheet

    def read(self, *args, **kwargs):
        return tuple(super(FakeTableSheet, self).read(*args, **kwargs))


data = (
    ("Col1Row1", "Col2Row1", "Col3Row1"),
    ("Col1Row2", "Col2Row2", "Col3Row2"),
    ("Col1Row3", "Col2Row3", "Col3Row3")
)


class TemplatedSheetTestCase(TestCase):
    def setUp(self):
        self.sheet = TestTemplatedSheet(sheetname="Testsheet")

    def test_columns(self):
        self.assertEqual(
            [self.sheet.column1, self.sheet.column2, self.sheet.column3],
            self.sheet.columns
        )

    def test_column_headers_not_unique_exception(self):
        class InvalidSheet(TableSheet):
            column1 = TableColumn(header="header")
            column2 = TableColumn(header="header")

        with self.assertRaises(ColumnHeadersNotUnique):
            InvalidSheet(sheetname="invalid_sheet")

    def test_column_headers_not_unique_multiple_conflicts(self):
        class InvalidSheet(TableSheet):
            column1 = TableColumn(header="header1")
            column2 = TableColumn(header="header1")
            column3 = TableColumn(header="header2")
            column4 = TableColumn(header="header2")

        with self.assertRaises(ColumnHeadersNotUnique):
            InvalidSheet(sheetname="invalid_sheet")

    def test_set_column_index(self):
        self.assertEqual(self.sheet.column1.column_index, 1)
        self.assertEqual(self.sheet.column2.column_index, 2)
        self.assertEqual(self.sheet.column3.column_index, 3)

    def test_no_table_columns_exeption(self):
        class NoColumnsTableSheet(TableSheet):
            pass

        with self.assertRaises(NoTableColumns):
            ws = NoColumnsTableSheet(sheetname="no_columns")

    def test_cannot_hide_or_group_last_column(self):
        class CannotHideLastColumnSheet(TableSheet):
            column1 = TableColumn()
            column2 = TableColumn(hidden=True)

        class CannotGroupLastColumnSheet(TableSheet):
            column1 = TableColumn()
            column2 = TableColumn(hidden=True)

        with self.assertRaises(CannotHideOrGroupLastColumn):
            CannotHideLastColumnSheet(sheetname="CannotHideOrGroupLastColumnSheet")

        with self.assertRaises(CannotHideOrGroupLastColumn):
            CannotGroupLastColumnSheet(sheetname="CannotGroupLastColumnSheet")

    def test_multiple_frozen_columns(self):
        class MultipleFrozenColumnsSheet(TableSheet):
            column1 = TableColumn(freeze=True)
            column2 = TableColumn(freeze=True)

        with self.assertRaises(MultipleFrozenColumns):
            MultipleFrozenColumnsSheet(sheetname="MultipleFrozenColumnsSheet")

    def test_read(self):
        obj = self.sheet.object_from_row(FakeCells("1", "2", "3"), row_number=3)
        self.assertEqual(obj.column1, "1")
        self.assertEqual(obj.column2, "2")
        self.assertEqual(obj.column3, "3")

    def test_find_headers_and_read(self):
        sheet = FakeTableSheet(
            ("column1", "column2", "column3"),
            ("1", "2", "3"),
        )

        objects = tuple(sheet.read())
        self.assertEqual(len(objects), 1)

    def test_headers_not_found(self):
        row_sets = (
            (
                ("1", "2", "3"),
                ("1", "2", "3"),
            ),
            (
                ("column1", "column2"),
                ("1", "2", "3"),
            )
        )

        for rows in row_sets:
            with self.assertRaises(HeadersNotFound):
                sheet = FakeTableSheet(rows)
                sheet.read()

    def test_no_rows(self):
        sheet = FakeTableSheet(
            (
                ("column1", "column2", "column3")
            ),
        )
        self.assertFalse(sheet.read())

    def test_write_tuple(self):
        wb = TestTemplatedWorkbook()

        wb.sheet1.write(objects=data)
        result = tuple(tuple(row) for row in wb.sheet1.read())

        self.assertEqual(data, result)

    def test_remove(self):
        wb = TestTemplatedWorkbook()
        self.assertTrue(wb.sheet1.empty)
        wb.sheet1.write(((1, 2, 3),))
        self.assertFalse(wb.sheet1.empty)
        wb.sheet1.remove()
        self.assertTrue(wb.sheet1.empty)

    def test_preserve(self):
        wb = TestTemplatedWorkbook()

        wb.sheet1.write(data[0:1])
        wb.sheet1.write(data[1:], preserve=True)

        result = tuple(tuple(row) for row in wb.sheet1.read())
        self.assertEqual(data, result)

    def test_do_not_preserve(self):
        wb = TestTemplatedWorkbook()

        wb.sheet1.write(data[0:1])
        wb.sheet1.write(data[1:], preserve=False)

        result = tuple(tuple(row) for row in wb.sheet1.read())
        self.assertEqual(data[1:], result)

    def test_no_freeze_pane(self):
        class NotFrozenWorkbook(TemplatedWorkbook):
            sheet1 = TestTemplatedSheet(freeze_header=False)

        wb = NotFrozenWorkbook()
        wb.sheet1.write(data)
        self.assertIsNone(wb.sheet1.worksheet.freeze_panes)

    def test_freeze_header(self):
        class FreezeHeaderWorkbook(TemplatedWorkbook):
            sheet1 = TestTemplatedSheet(freeze_header=True)
            sheet2 = TestTemplatedSheet(freeze_header=True)

        wb = FreezeHeaderWorkbook()
        wb.sheet1.write(data)

        self.assertEqual(wb.sheet1.worksheet.freeze_panes, "A2")

        wb.sheet2.write(data, title="Title")
        self.assertEqual(wb.sheet2.worksheet.freeze_panes, "A3")

    def test_freeze_column(self):
        class FreezeFirstSheet(TableSheet):
            column1 = TableColumn(header="column1", freeze=True)
            column2 = TableColumn(header="column2")
            column3 = TableColumn(header="column3")

        class FreezeSecondSheet(TableSheet):
            column1 = TableColumn(header="column1")
            column2 = TableColumn(header="column2", freeze=True)
            column3 = TableColumn(header="column3")

        class FreezeThirdSheet(TableSheet):
            column1 = TableColumn(header="column1")
            column2 = TableColumn(header="column2")
            column3 = TableColumn(header="column3", freeze=True)

        class FrozenColumnWorkbook(TemplatedWorkbook):
            sheet1 = FreezeFirstSheet(freeze_header=False)
            sheet2 = FreezeSecondSheet(freeze_header=False)
            sheet3 = FreezeThirdSheet(freeze_header=False)

        wb = FrozenColumnWorkbook()
        for sheet, cell in ((wb.sheet1, "B1"), (wb.sheet2, "C1"), (wb.sheet3, "D1")):
            sheet.write(data)
            self.assertEqual(cell, sheet.worksheet.freeze_panes)

        class FrozenColumnAndHeaderWorkbook(TemplatedWorkbook):
            sheet1 = FreezeFirstSheet(freeze_header=True)
            sheet2 = FreezeSecondSheet(freeze_header=True)
            sheet3 = FreezeThirdSheet(freeze_header=True)

        wb = FrozenColumnAndHeaderWorkbook()
        for sheet, cell in ((wb.sheet1, "B2"), (wb.sheet2, "C2"), (wb.sheet3, "D2")):
            sheet.write(data)
            self.assertEqual(cell, sheet.worksheet.freeze_panes)
