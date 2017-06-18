from collections import OrderedDict
from unittest import TestCase

from openpyxl_templates.table_sheet.columns import TableColumn, ColumnIndexNotSet
from openpyxl_templates.table_sheet.sheet import TableSheet, ColumnHeadersNotUnique, NoTableColumns
from openpyxl_templates.templated_workbook import TemplatedWorkbook
from openpyxl_templates.utils import OrderedType, class_property


class MagicString(str):
    pass


class OrderedTypeTestClass(metaclass=OrderedType):
    item_class = MagicString

    @class_property
    def items(self):
        return list(self._items.values())


class OrderedTypeTests(TestCase):
    def test_objects_identified(self):
        class Test(OrderedTypeTestClass):
            string1 = MagicString("string1")
            string2 = MagicString("string2")
            string3 = MagicString("string3")

        result = list(Test.items)
        for index, string in enumerate((Test.string1, Test.string2, Test.string3)):
            self.assertEqual(result[index], string)

    def test_inheritence(self):
        class Test(OrderedTypeTestClass):
            string1 = MagicString("string1")
            string2 = MagicString("string2")
            string3 = MagicString("string3")

        class Test2(Test):
            string2 = MagicString("new_string2")
            string4 = MagicString("string4")

        result = list(Test2.items)
        for index, string in enumerate((Test2.string1, Test2.string2, Test2.string3, Test2.string4)):
            self.assertEqual(result[index], string)

    def test_multiple_inheritence(self):
        class Parent1(OrderedTypeTestClass):
            string1 = MagicString("Parent1.string1")
            string2 = MagicString("Parent1.string2")

        class Parent2(OrderedTypeTestClass):
            string2 = MagicString("Parent2.string2")
            string3 = MagicString("Parend2.string3")

        class Child1(Parent2, Parent1):
            string3 = MagicString("child.string3")
            string4 = MagicString("child.string4")

        class Child2(Parent1, Parent2):
            string3 = MagicString("child.string3")
            string4 = MagicString("child.string4")

        result = Child1.items
        for index, attr in enumerate(["string1", "string2", "string3", "string4"]):
            self.assertEqual(result[index], getattr(Child1, attr))

        result = Child2.items
        for index, attr in enumerate(["string2", "string3", "string1", "string4"]):
            self.assertEqual(result[index], getattr(Child2, attr))


class TestColumn(TableColumn):
    _header = "Test"


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


class TestTemplatedSheet(TableSheet):
    column1 = TestColumn(header="column1")
    column2 = TestColumn(header="column2")
    column3 = TestColumn(header="column3")


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
            column1 = TestColumn(header="header")
            column2 = TestColumn(header="header")

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




class TestTemplatedWorkbook(TemplatedWorkbook):
    sheet1 = TestTemplatedSheet(sheetname="Test")
    sheet2 = TestTemplatedSheet(sheetname="Test2")


class InheritingTemplatedWorkbook(TestTemplatedWorkbook):
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




