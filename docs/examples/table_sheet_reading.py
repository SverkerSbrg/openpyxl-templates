from openpyxl_templates import TemplatedWorksheet
from openpyxl_templates.table_sheet import TableSheet, CharColumn, IntColumn


class DemoTableSheet(TableSheet):
    column1 = CharColumn()
    column2 = IntColumn()

class DemoTemplatedWorksheet(TemplatedWorksheet):
    demo_sheet = DemoTableSheet()

wb = DemoTableSheet()

# TODO: Continue when writing is done