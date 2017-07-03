from itertools import repeat
from openpyxl_templates import TemplatedWorkbook
from openpyxl_templates.table_sheet import TableSheet
from openpyxl_templates.table_sheet.columns import CharColumn


class TableSheetElements(TableSheet):
    column1 = CharColumn(header="Header 1")
    column2 = CharColumn(header="Header 2")
    column3 = CharColumn(header="Header 3")
    column4 = CharColumn(header="Header 4")


class TableSheetElementsWorkook(TemplatedWorkbook):
    table_sheet_elements = TableSheetElements()


wb = TableSheetElementsWorkook()
wb.table_sheet_elements.write(
    title="Title",
    description="This is the description, it can be a couple of sentences long.",
    objects=(tuple(repeat("Row %d" % i, times=4)) for i in range(1, 4))
)
wb.save("table_sheet_elements.xlsx")
