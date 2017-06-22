from unittest import TestCase

from openpyxl_templates.table_sheet.columns import TableColumn
from openpyxl_templates.table_sheet.sheet import TableSheet, ColumnHeadersNotUnique, NoTableColumns, \
    CannotHideOrGroupLastColumn, HeadersNotFound
from tests.utils import FakeCells


class TestTemplatedSheet(TableSheet):
    column1 = TableColumn(header="column1")
    column2 = TableColumn(header="column2")
    column3 = TableColumn(header="column3")


class FakeTableSheet(TableSheet):
    column1 = TableColumn(header="column1")
    column2 = TableColumn(header="column2")
    column3 = TableColumn(header="column3")

    def __init__(self, *rows):
        self.fake_worksheet = (FakeCells(*row) for row in rows)

        super().__init__(sheetname="fakesheet")

    @property
    def worksheet(self):
        return self.fake_worksheet

    def read(self, *args, **kwargs):
        return tuple(super().read(*args, **kwargs))


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

    def test_read(self):
        obj = self.sheet.object_from_row(FakeCells("1", "2", "3"))
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
