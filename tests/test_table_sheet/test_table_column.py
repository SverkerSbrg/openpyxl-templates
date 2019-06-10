from datetime import datetime
from unittest import TestCase

from openpyxl import Workbook

from openpyxl_templates import TemplatedWorkbook
from openpyxl_templates.table_sheet import TableSheet
from openpyxl_templates.table_sheet.columns import TableColumn, ColumnIndexNotSet, BoolColumn, StringToLong, \
    CharColumn, UnableToParseBool, FloatColumn, BlankNotAllowed, UnableToParseFloat, IntColumn, RoundingRequired, \
    ChoiceColumn, IllegalChoice, DatetimeColumn, UnableToParseDatetime
from openpyxl_templates.utils import FakeCell


class ColumnTestCase(TestCase):
    def assertToExcel(self, excel, internal, column=None):
        column = column or self.column
        result = column._to_excel(internal)
        self.assertEqual(
            result, 
            excel,
            msg="Internal value '%s' (%s) was converted to '%s' (%s) instead of '%s' (%s)." % (
                internal, type(internal).__name__,
                result, type(result).__name__,
                excel, type(excel).__name__
            )
        )
        
    def assertFromExcel(self, excel, internal, column=None):
        column = column or self.column
        result = column._from_excel(FakeCell(excel))
        self.assertEqual(
            result, 
            internal,
            msg="Excel value '%s' (%s) was converted to '%s' (%s) instead of '%s' (%s)." % (
                excel, type(excel).__name__,
                result, type(result).__name__,
                internal, type(internal).__name__
            )
        )


class TableColumnTests(ColumnTestCase):
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

    def test_from_excel_None(self):
        for value in ("", None, "'"):
            self.assertFromExcel(value, None)

    def test_blank_not_allowed(self):
        column = TableColumn(allow_blank=False)

        for value in ("", None):
            with self.assertRaises(BlankNotAllowed, msg="'%s' (%s)" % (value, type(value).__name__)):
                column._to_excel(value)

        for value in ("", None, "'"):
            with self.assertRaises(BlankNotAllowed, msg="'%s' (%s)" % (value, type(value).__name__)):
                column._from_excel(FakeCell(value))


class CharColumnTests(ColumnTestCase):
    def setUp(self):
        self.column = CharColumn()

    def test_to_excel(self):
        for excel, internal in (
                (None, None),
                ("1", 1),
                ("1.0", 1.0),
                ("String", "String"),
                (None, "")
        ):
            self.assertToExcel(excel, internal)

    def test_from_excel(self):
        for excel, internal in (
                (None, None),
                (1, "1"),
                (1.0, "1.0"),
                ("String", "String"),
                ("", None)
        ):
            self.assertFromExcel(excel, internal)

    def test_string_to_long(self):
        column = CharColumn(max_length=5)
        column._from_excel(FakeCell("12345"))

        with self.assertRaises(StringToLong):
            column._from_excel(FakeCell("123456"))


class BooleanColumnTests(ColumnTestCase):
    def setUp(self):
        self.column = BoolColumn()

    def test_to_excel(self):
        for excel, internal in (
                # ("TRUE", True),
                # ("FALSE", False),
                (None, None),
                (None, ""),
                # ("TRUE", "string")
        ):
            self.assertToExcel(excel, internal)

    def test_from_excel(self):
        for excel, internal in (
                ("TRUE", True),
                # ("FALSE", False),
                ("'TRUE", True),
                # ("'FALSE", False),
                (1, True),
                (0, False),
                ("x", True),
                ("", None)
        ):
            self.assertFromExcel(excel, internal)

    def test_strict(self):
        column = BoolColumn(strict=True)

        for valid in (True, False):
            column._from_excel(FakeCell(valid))

        for invalid in ("string",):
            with self.assertRaises(UnableToParseBool, msg=str(invalid)):
                column._from_excel(FakeCell(invalid))

    def test_modify_excel_true_false(self):
        column = BoolColumn(
            excel_false="F",
            excel_true="T"
        )
        for excel, internal in (
                ("T", True),
                ("F", False),
        ):
            self.assertFromExcel(excel, internal, column=column)
            self.assertToExcel(excel, internal, column=column)

    def test_create_cell_only_use_default_on_None(self):
        column = BoolColumn(object_attribute="test", row_style="")
        self.assertEqual(column.create_cell(Workbook().active, False).value, column.excel_false)


class FloatColumnTestCase(ColumnTestCase):
    def setUp(self):
        self.column = FloatColumn()

    def test_default_kwargs(self):
        self.assertEqual(self.column.cell_style, "Row, decimal")
        self.assertEqual(self.column.default, 0.0)

    def test_override_default_kwargs(self):
        column = FloatColumn(default=1.3, row_style="New style")

        self.assertEqual(column.cell_style, "New style")
        self.assertEqual(column.default, 1.3)

    def test_to_excel(self):
        for excel, internal in (
                (1, 1.0),
                (1.3, 1.3),
                (-1.1, -1.1),
                (0, 0),
                (self.column.default, None),
                (self.column.default, "")
        ):
            self.assertToExcel(excel, internal)

    def test_from_excel(self):
        for excel, internal in (
                (1, 1.0),
                (1.1, 1.1),
                (-1.0, -1.0),
                ("1.3", 1.3),
                ("-1.0", -1.0),
                ("20", 20),
                ("'3", 3),
                ("'", 0),
                ("    0.1   ", 0.1),
                (None, 0)
        ):
            self.assertFromExcel(excel, internal)

    def test_unable_to_parse_float(self):
        for value in ("String", object()):
            with self.assertRaises(UnableToParseFloat, msg=value):
                self.column._to_excel(value)

        for value in ("String",):
            with self.assertRaises(UnableToParseFloat, msg=value):
                self.column._from_excel(FakeCell(value))

class TestCaseIntColumn(ColumnTestCase):
    def setUp(self):
        self.column = IntColumn()

    def test_to_excel(self):
        for excel, internal in (
                (1, 1.0),
                (1, 1.1),
                (-1, -1.1),
                (2, 1.9),
                (0, 0),
                (self.column.default, None),
                (self.column.default, "")
        ):
            self.assertToExcel(excel, internal)

    def test_from_excel(self):
        for excel, internal in (
                (1, 1),
                (1.1, 1),
                (-1.0, -1),
                (1.9, 2),
                ("1.3", 1),
                ("1.9", 2),
                ("-1.0", -1),
                ("20", 20),
                ("'3", 3),
                ("'", 0),
                (None, 0)
        ):
            self.assertFromExcel(excel, internal)

    def test_rounding_required(self):
        column = IntColumn(round_value=False)

        for value in ("1.3", 1.3, 0.99999999, "-23.6"):
            with self.assertRaises(RoundingRequired):
                column._to_excel(value)

            with self.assertRaises(RoundingRequired):
                column._from_excel(FakeCell(value))

        # Does not raise
        for value in (1, 1.0, -0.0, None):
            column._to_excel(value)
            column._from_excel(FakeCell(value))


class ChoiceColumnTestCase(ColumnTestCase):
    def setUp(self):
        self.choices = (
            ("internal1", "excel1"),
            ("internal2", "excel2"),
            ("internal3", "excel3"),
        )
        self.column = ChoiceColumn(choices=self.choices)

    def test_to_excel(self):
        for excel, internal in (
                ("excel1", "internal1"),
                ("excel2", "internal2"),
                ("excel3", "internal3"),
                (None, None),
        ):
            self.assertToExcel(excel, internal)

    def test_from_excel(self):
        for excel, internal in (
                ("excel1", "internal1"),
                ("excel2", "internal2"),
                ("excel3", "internal3"),
                (None, None),
        ):
            self.assertFromExcel(excel, internal)

    def test_illegal_choice(self):
        for value in (1, "string"):
            with self.assertRaises(IllegalChoice):
                self.column._from_excel(FakeCell(value))

            with self.assertRaises(IllegalChoice):
                self.column._to_excel(value)

    def test_default_value(self):
        column = ChoiceColumn(choices=self.choices, default="internal1")

        for value in (None, "String", 1):
            self.assertFromExcel(value, "internal1", column=column)

        for value in (None, "String", 1):
            self.assertToExcel("excel1", value, column=column)

    def test_illegal_default(self):
        with self.assertRaises(IllegalChoice):
            column = ChoiceColumn(choices=self.choices, default="not_a_choice")

    def test_choices_required(self):
        with self.assertRaises(ValueError):
            column = ChoiceColumn()

    def test_choices_as_generator(self):
        column = ChoiceColumn(
            choices=((value, value) for value in ("value1", "value2"))
        )
        self.assertEqual(
            list(column.choices),
            [("value1", "value1"), ("value2", "value2")]
        )


class DatetimeColumnTestCase(TableColumnTests):
    def setUp(self):
        self.column = DatetimeColumn()

    def test_to_excel(self):
        for excel, internal in (
                (1.0, datetime(year=1900, month=1, day=1)),
                (1.5, datetime(year=1900, month=1, day=1, hour=12)),
                (59, datetime(year=1900, month=2, day=28)),
                (61, datetime(year=1900, month=3, day=1)),
                (43006, datetime(year=2017, month=9, day=28)),
                (43006.628958333335, datetime(year=2017, month=9, day=28, hour=15, minute=5, second=42)),
        ):
            self.assertToExcel(excel, internal)

    def test_from_excel(self):
        for excel, internal in (
                (1.0, datetime(year=1900, month=1, day=1)),
                (1.5, datetime(year=1900, month=1, day=1, hour=12)),
                (59, datetime(year=1900, month=2, day=28)),
                (61, datetime(year=1900, month=3, day=1)),
                (43006, datetime(year=2017, month=9, day=28)),
                (43006.628958333335, datetime(year=2017, month=9, day=28, hour=15, minute=5, second=42)),
                (
                    datetime(year=2017, month=9, day=28, hour=15, minute=5, second=42),
                    datetime(year=2017, month=9, day=28, hour=15, minute=5, second=42)
                )
        ):
            self.assertFromExcel(excel, internal)

    def test_unable_to_parse_datetime(self):
        for value in ("2017-09-28", -1, -1000.0, "adsf", False, object()):
            with self.assertRaises(UnableToParseDatetime, msg=value):
                self.column._from_excel(FakeCell(value))

            with self.assertRaises(UnableToParseDatetime, msg=value):
                self.column._to_excel(value)

