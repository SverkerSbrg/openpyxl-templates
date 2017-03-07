from unittest import TestCase

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font
from openpyxl.styles import NamedStyle

from tests.style2 import NamedStyleManager

#
# class StylesTestCase(TestCase):
#     def setUp(self):
#         self.workbook = Workbook()
#         self.worksheet = self.workbook.active
#         # self.bold1 = NamedStyle(name="bold", font=Font(bold=True))
#         # self.bold2 = NamedStyle(name="bold", font=Font(bold=True))
#
#     def test_equals(self):
#         style1 = NamedStyle(font=Font(bold=True, size=12))
#         style2 = NamedStyle(font=Font(bold=False, italic=True))
#
#         self.assertEqual(
#             NamedStyle("Result", font=Font(italic=True, size=12)),
#             NamedStyleManager.merge(style1, style2, result_name="Result")
#         )
#
#     def test_no_duplicates(self):
#         style1 = NamedStyle(font=Font(bold=True, size=12))
#         style2 = NamedStyle(font=Font(bold=True, size=12))
#
#         self.assertIs(style1, NamedStyleManager.merge(style1, style2, result_name="Does not matter"))
#         self.assertIs(style2, NamedStyleManager.merge(None, style2, style1, None, result_name="Does not matter either"))
#
#     # def test_no_auto_clear_bold(self):
    #     style1 = NamedStyle(font=Font(bold=True, size=12))
    #     style2 = NamedStyle(font=Font(italic=True))
    #
    #     self.assertEqual(
    #         NamedStyle(font=Font(italic=True, bold=True, size=12)),
    #         NamedStyleManager.merge(style1, style2)
    #     )

