from openpyxl import Workbook, load_workbook

from openpyxl_templates.style import StyleSet, StandardStyleSet, DefaultStyleSet, StyleSet2
from openpyxl_templates.templated_sheet import TemplatedSheet
from openpyxl_templates.utils import OrderedType, Typed


class TemplatedWorkbook(Workbook, metaclass=OrderedType):
    item_class = TemplatedSheet

    templated_sheets = None
    template_styles = Typed("template_styles", expected_type=StyleSet2)

    def __new__(cls, *args, file=None, **kwargs):
        if file:
            return load_workbook(file)
        return super().__new__(cls)

    def __init__(self, template_styles=None):
        super().__init__()

        self.templated_sheets = list(self._items.values())
        for sheet in self.templated_sheets:
            sheet.workbook = self
        self.template_styles = template_styles or self.template_styles or DefaultStyleSet()

    def remove_all_sheets(self):
        for sheetname in self.sheetnames:
            del self[sheetname]
