from openpyxl_templates.exceptions import OpenpyxlTemplateException
from openpyxl_templates.utils import OrderedType, Typed


class TemplatedWorkbookNotSet(OpenpyxlTemplateException):
    def __init__(self, templated_sheet):
        super().__init__(
            "The sheet '%s' has no assosiated workbook. This should be done automatically by the TemplatedWorkbook."
            % templated_sheet.sheetname
        )


class WorksheetDoesNotExist(OpenpyxlTemplateException):
    def __init__(self, templated_sheet):
        super().__init__(
            "The workbook has no sheet '%s'." % templated_sheet.sheetname
        )


class TemplatedSheet(metaclass=OrderedType):
    sheetname = Typed("sheetname", expected_type=str)
    active = Typed("active", expected_type=bool, value=False)
    _workbook = None

    # order = ... # TODO: Add ordering to sheets either through declaration on workbook or here

    def __init__(self, sheetname=None, active=None):
        self.sheetname = sheetname or self.sheetname
        self.active = active if active is not None else self.active

    @property
    def exists(self):
        return self.sheetname in self.workbook

    @property
    def worksheet(self):
        if not self.exists:
            self.workbook.create_sheet(self.sheetname)

        return self.workbook[self.sheetname]

    @property
    def workbook(self):
        if not self._workbook:
            raise TemplatedWorkbookNotSet(self)
        return self._workbook

    @workbook.setter
    def workbook(self, workbook):
        self._workbook = workbook

    @property
    def sheet_index(self):
        try:
            return self.workbook.sheetnames.index(self.sheetname)
        except ValueError:
            raise WorksheetDoesNotExist(self)

    def write(self, *args, overwrite=True, **kwargs):
        raise NotImplemented()
        # 'self.sheet_template.write(self.worksheet, self.templated_workbook.styles, data)

    def read(self, exception_policy):
        raise NotImplemented()

    def remove(self):
        if self.exists:
            del self.workbook[self.sheetname]

    def activate(self):
        self.workbook.active = self.sheet_index
