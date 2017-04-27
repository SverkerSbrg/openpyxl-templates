from unittest import TestCase

from openpyxl import Workbook

from tests.demo import DemoTemplate


class WriteTestCase(TestCase):
    def test_write_empty(self):
        workbook = Workbook()
        template = DemoTemplate(workbook)
        template.remove_all_sheets()
        template.write_sheet("Persons", [])