from unittest import TestCase

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font
from openpyxl.styles import NamedStyle


class StylesTestCase(TestCase):
    def setUp(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.bold1 = NamedStyle(name="bold", font=Font(bold=True))
        self.bold2 = NamedStyle(name="bold", font=Font(bold=True))

    def test_equals(self):
        bold1 = WriteOnlyCell(ws=self.worksheet)
        bold1.style = self.bold1
        bold2 = WriteOnlyCell(ws=self.worksheet)
        bold2.style = self.bold2
        self.worksheet.append([bold1])
        self.worksheet.append([bold2])
        print(self.workbook.named_styles)
