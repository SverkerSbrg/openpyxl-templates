from openpyxl_templates.table_sheet import TableSheet
from openpyxl_templates.table_sheet.columns import CharColumn, IntColumn


class DemoTableSheet(TableSheet):
    column1 = CharColumn()
    column2 = IntColumn()


ws = DemoTableSheet()
assert (tuple(ws.columns) == (ws.column1, ws.column2))


# ------------------- Inheritance of TableSheet -------------------
class ExtendedDemoTableSheet(DemoTableSheet):
    column3 = CharColumn()

ws = ExtendedDemoTableSheet()
assert (tuple(ws.columns) == (ws.column1, ws.column2, ws.column3))


# ------------------- Automatic heading -------------------
class DemoTableSheet(TableSheet):
    column1 = CharColumn(header="Header 1")
    column2 = IntColumn()  # The header of column2 will be set automatically to "column2"



