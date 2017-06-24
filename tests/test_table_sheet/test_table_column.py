from unittest import TestCase

from openpyxl_templates.table_sheet.columns import TableColumn, ColumnIndexNotSet, BoolColumn, StringToLong, \
    CharColumn, UnableToParseBool
from tests.utils import FakeCell


class TableColumnTests(TestCase):
    def setUp(self):
        self.column = TableColumn(header="TestColumn")

    def test_column_index_not_set_exception(self):
        with self.assertRaises(ColumnIndexNotSet):
            i = self.column.column_index

    def test_auto_column_header(self):
        column = TableColumn()
        column.column_index = 1

        self.assertFalse(column._header)
        header = column.header
        self.assertIsInstance(header, str)
        self.assertTrue(header)

    def test_column_letter(self):
        column = TableColumn()

        for i in range(1, 20):
            column.column_index = i
            self.assertEqual(
                column.column_letter,
                chr(ord('A') + i - 1)
            )



class CharColumnTests(TestCase):
    def setUp(self):
        self.column = CharColumn()

    def test_to_excel(self):
        for excel, internal in (
                ("", None),
                ("1", 1),
                ("1.0", 1.0),
                ("String", "String"),
                ("", "")
        ):
            self.assertEqual(self.column.to_excel(internal), excel)

    def test_from_excel(self):
        for excel, internal in (
                (None, None),
                (1, "1"),
                (1.0, "1.0"),
                ("String", "String"),
                ("", "")
        ):
            self.assertEqual(self.column.from_excel(FakeCell(excel)), internal)

    def test_string_to_long(self):
        column = CharColumn(max_length=5)
        column.from_excel(FakeCell("12345"))

        with self.assertRaises(StringToLong):
            column.from_excel(FakeCell("123456"))


class BooleanColumnTests(TestCase):
    def setUp(self):
        self.column = BoolColumn()

    def test_to_excel(self):
        for excel, internal in (
                ("TRUE", True),
                ("FALSE", False),
                ("FALSE", None),
                ("FALSE", ""),
                ("TRUE", "string")
        ):
            self.assertEqual(self.column.to_excel(internal), excel)

    def test_from_excel(self):
        for excel, internal in (
                ("TRUE", True),
                ("FALSE", False),
                (1, True),
                (0, False),
                ("x", True),
                ("", False)
        ):
            self.assertEqual(self.column.from_excel(FakeCell(excel)), internal)

    def test_strict(self):
        column = BoolColumn(strict=True)

        for valid in ("TRUE", "FALSE", True, False):
            column.from_excel(FakeCell(valid))

        for invalid in (None, "string", "", 0, 1, 0.1):
            with self.assertRaises(UnableToParseBool):
                column.from_excel(FakeCell(invalid))
