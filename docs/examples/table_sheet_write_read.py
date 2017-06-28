from openpyxl_templates import TemplatedWorkbook
from openpyxl_templates.table_sheet import TableSheet, CharColumn, IntColumn


class DemoTableSheet(TableSheet):
    column1 = CharColumn()
    column2 = IntColumn()


class DemoTemplatedWorksheet(TemplatedWorkbook):
    demo_sheet1 = DemoTableSheet()
    demo_sheet2 = DemoTableSheet()

wb = DemoTemplatedWorksheet()

wb.demo_sheet1.write(
    objects=(
        ("Row 1", 1),
        ("Row 2", 2),
        ("Row 3", 3),
    ),
    title="The first sheet"
)
wb.demo_sheet2.write(
    objects=(
        ("Row 1", 1),
        ("Row 2", 2),
        ("Row 3", 3),
    ),
    title="The second sheet",
    description="Lorem ipsum dolor sit amet, consectetur adipiscing elit. In euismod, sem eu."
)
wb.save("read_write.xlsx")


# --------------- Passing Objects ---------------

class DemoObject():
    def __init__(self, column1, column2):
        self.column1 = column1
        self.column2 = column2


wb = DemoTemplatedWorksheet()
wb.demo_sheet1.write(
    objects=(
        DemoObject("Row 1", 1),
        DemoObject("Row 2", 2),
        DemoObject("Row 3", 3),
    )
)


# ------------------- Styling -------------------
class PrettyDemoSheet(TableSheet):
    title_style = "Bold & red, title"
    description_style = "Extra tiny, description"

    column1 = CharColumn()
    column2 = IntColumn()


# --------------- Reading ---------------
wb = DemoTemplatedWorksheet("read_write.xlsx")
for row in wb.demo_sheet1.read():
    print(row)


# --------------- As iterator ---------------
for row in wb.demo_sheet2:
    print(row)

