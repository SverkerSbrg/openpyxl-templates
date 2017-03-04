from style import CellStyle


class WorkbookTemplate:
    sheets = None
    active_sheet = None

    style = None
    header_style = None
    title_style = None
    description_style = None

    def __init__(self, workbook, sheets=None, active_sheet=None, style=None, header_style=None, title_style=None, description_style=None):
        self.workbook = workbook
        self.sheets = sheets or self.sheets or []
        self.active_sheet = active_sheet or self.active_sheet

        self.style = CellStyle.merge(self.style, style)
        self.header_style = CellStyle.merge(self.style, self.header_style, header_style)
        self.title_style = CellStyle.merge(self.style, self.title_style, title_style)
        self.description_style = CellStyle.merge(self.style, self.description_style, description_style)

        for sheet in self.sheets:
            sheet.style = CellStyle.merge(self.style, sheet.style)
            sheet.header_style = CellStyle.merge(self.style, sheet.header_style)
            sheet.title_style = CellStyle.merge(self.style, sheet.title_style)
            sheet.description_style = CellStyle.merge(self.style, sheet.description_style)
            sheet.rebase_column_styles()

        self._sheet_map = {sheet.name: sheet for sheet in self.sheets}

    def write_sheet(self, name, objects):
        excel_sheet = self._sheet_map[name]
        worksheet = self.get_or_create_sheet(excel_sheet)
        excel_sheet.write(worksheet, objects)

        self.update_active_sheet()

    def read_rows(self, name):
        excel_sheet = self._sheet_map[name]
        worksheet = self.get_or_create_sheet(excel_sheet)
        return excel_sheet.read_rows(worksheet)

    def get_or_create_sheet(self, excel_sheet):
        name = excel_sheet.name
        if name in self.workbook.sheetnames:
            return self.workbook[name]
        return self.workbook.create_sheet(excel_sheet.name)

    def update_active_sheet(self):
        for index, sheetname in enumerate(self.workbook.sheetnames):
            if sheetname == self.active_sheet:
                self.workbook.active = index
                return

