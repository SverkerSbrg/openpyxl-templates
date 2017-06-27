from openpyxl_templates.table_sheet import TableSheet, CharColumn, IntColumn


class DemoTableSheet(TableSheet):
    column1 = CharColumn()
    column2 = IntColumn()


ws = DemoTableSheet()
assert (tuple(ws.columns) == (ws.column1, ws.column2))


class InheritingTableSheet(DemoTableSheet):
    column3 = CharColumn()


class DemoTableSheet(TableSheet):
    column1 = CharColumn(header="Header 1")
    column2 = IntColumn()  # The header of column2 will be set automatically to "column2"
