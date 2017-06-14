from collections import OrderedDict

from openpyxl_templates.columns import Column
from openpyxl_templates.utils import Typed
from openpyxl_templates.workbook import TemplatedWorkbook


class OrderedType(type):
    @classmethod
    def __prepare__(mcs, name, bases):
        return OrderedDict()

    def __new__(mcs, name, bases, classdict):
        obj = super().__new__(mcs, name, bases, classdict)
        obj.__register_objects__(obj, classdict)

    def __register_objects__(cls, classdict):
        pass

class Test(metaclass=OrderedType):
    x1 = 2
    x2 = 3
    x3 = 4

    # def __new__(cls, *args, **kwargs):
    #     print(args, kwargs)
    #     return super().__new__(cls, *args, **kwargs)

    def __register_objects__(self, classdict):
        print(classdict)




class TemplatedSheet(metaclass=OrderedType):
    sheetname = Typed("sheetname", expected_type=str)
    active = Typed("active", expected_type=bool, value=False)
    workbook = Typed("workbook", expected_type=TemplatedWorkbook, allow_none=True)

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

    def write(self, *args, overwrite=True, **kwargs):
        raise NotImplemented()
        # 'self.sheet_template.write(self.worksheet, self.templated_workbook.styles, data)

    def read(self, exception_policy):
        raise NotImplemented()

    def remove(self):
        if self.exists:
            del self.workbook[self.sheetname]

    def activate(self):
        self.workbook.active = self.worksheet


class TableSheet(TemplatedSheet):
    _columns = None

    freeze_header = Typed("freeze_header", expected_type=bool, value=True)
    hide_excess_columns = Typed("hide_excess_columns", expected_type=bool, value=True)
    format_as_table = Typed("format_as_table", expected_type=bool, value=True)

    def __register_objects__(self, classdict):
        self._columns = list(value for value in classdict.values if issubclass(type(value), Column))

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def write(self, title=None, description=None, objects=None):
        pass

    def read(self, exception_policy):
        pass

    @property
    def columns(self):
        return self._columns




