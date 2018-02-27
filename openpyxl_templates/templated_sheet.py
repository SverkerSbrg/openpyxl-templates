from future.utils import with_metaclass

from openpyxl_templates.exceptions import OpenpyxlTemplateException
from openpyxl_templates.utils import OrderedType, Typed


class TemplatedWorkbookNotSet(OpenpyxlTemplateException):
    def __init__(self, templated_sheet):
        super(TemplatedWorkbookNotSet, self).__init__(
            "The sheet '%s' has no assosiated workbook. This should be done automatically by the TemplatedWorkbook."
            % templated_sheet.sheetname
        )


class WorksheetDoesNotExist(OpenpyxlTemplateException):
    def __init__(self, templated_sheet):
        super(WorksheetDoesNotExist, self).__init__(
            "The workbook has no sheet '%s'." % templated_sheet.sheetname
        )


class SheetnameNotSet(OpenpyxlTemplateException):
    def __init__(self):
        super(SheetnameNotSet, self).__init__(
            "Sheetname not specified. This should be done automatically by the TemplatedWorkbook.")


class TemplatedWorksheet(with_metaclass(OrderedType)):
    _sheetname = Typed("_sheetname", expected_type=str, allow_none=True)
    active = Typed("active", expected_type=bool, value=False)
    _workbook = None
    template_styles = None

    # order = ... # TODO: Add ordering to sheets either through declaration on workbook or here

    def __init__(self, sheetname=None, active=None):
        self._sheetname = sheetname or self._sheetname
        self.active = active if active is not None else self.active

    @property
    def exists(self):
        return self.sheetname in self.workbook

    @property
    def empty(self):
        if not self.exists:
            return True

        return not bool(len(self.worksheet._cells))

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

    def write(self, data):
        raise NotImplemented()
        # 'self.sheet_template.write(self.worksheet, self.templated_workbook.styles, data)

    def read(self):
        raise NotImplemented()

    def remove(self):
        if self.exists:
            del self.workbook[self.sheetname]

    # def activate(self):
    #     self.workbook.active = self.sheet_index

    @property
    def sheetname(self):
        if not self._sheetname:
            raise SheetnameNotSet()
        return self._sheetname

    @sheetname.setter
    def sheetname(self, value):
        self._sheetname = value

    def __str__(self):
        return self._sheetname or self.__class__.__name__

    def __repr__(self):
        return str(self)
