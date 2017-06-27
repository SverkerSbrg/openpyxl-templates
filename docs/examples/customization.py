from openpyxl_templates import TemplatedWorkbook
from openpyxl_templates.table_sheet import TableSheet, CharColumn, TextColumn, FloatColumn


class IceCream:
    def __init__(self, name, description, flavour, color, price):
        self.name = name
        self.description = description
        self.flavour = flavour
        self.color = color
        self.price = price


class IceCreamSheet(TableSheet):
    name = CharColumn()
    description = TextColumn(width=32)
    flavour = CharColumn()
    color = CharColumn()
    price = FloatColumn()

    def create_object(self, data):
        return IceCream(**dict(data))


class IceCreamWorkbook(TemplatedWorkbook):
    ice_creams = IceCreamSheet()


wb = IceCreamWorkbook()
wb.ice_creams.write(
    title="Ice cream",
    objects=(("Tripple chocolate", "Really good chocolate ice cream.", "Chocolate", "Brown", 5.99),)
)
wb.save("customization.xlsx")

wb = IceCreamWorkbook("customization.xlsx")
for ice_cream in wb.ice_creams:
    assert isinstance(ice_cream, IceCream)
    print(ice_cream.name)
