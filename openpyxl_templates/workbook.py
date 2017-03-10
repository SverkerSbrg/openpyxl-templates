from openpyxl_templates.style import StyleSet, SheetStyleMixin, StandardStyleSet
from openpyxl_templates.utils import Typed


class WorkbookTemplate(SheetStyleMixin):
    sheets = None
    active_sheet = Typed("active_sheet", expected_type=str, allow_none=True)

    styles = Typed("styles", expected_type=StyleSet)

    header_style = "header"

    def __init__(self, workbook, sheets=None, active_sheet=None, styles=None, **style_keys):
        super().__init__(**style_keys)

        self.workbook = workbook
        self.sheets = sheets or self.sheets or []
        self.active_sheet = active_sheet or self.active_sheet

        self.styles = styles or self.styles or StandardStyleSet()

        self._sheet_map = {sheet.sheetname: sheet for sheet in self.sheets}

        for sheet in self.sheets:
            sheet.inherit_styles(self)

    def write_sheet(self, name, objects):
        excel_sheet = self._sheet_map[name]
        worksheet = self.get_or_create_sheet(excel_sheet)
        excel_sheet.write(worksheet, self.styles, objects)

        self.update_active_sheet()

    def read_rows(self, name):
        excel_sheet = self._sheet_map[name]
        worksheet = self.get_or_create_sheet(excel_sheet)
        return excel_sheet.read_rows(worksheet)

    def get_or_create_sheet(self, excel_sheet):
        name = excel_sheet.sheetname
        if name in self.workbook.sheetnames:
            return self.workbook[name]
        return self.workbook.create_sheet(excel_sheet.sheetname)

    def update_active_sheet(self):
        for index, sheetname in enumerate(self.workbook.sheetnames):
            if sheetname == self.active_sheet:
                self.workbook.active = index
                return
