from collections import OrderedDict
from types import FunctionType

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.columns import Column
from openpyxl_templates.exceptions import BlankNotAllowed
from openpyxl_templates.style import StyleSet, StandardStyleSet
from openpyxl_templates.utils import Typed, OrderedType


class Test():
    t = 2

    def __new__(cls, *args, **kwargs):
        print(args, kwargs)

        super().__new__(cls, *args, **kwargs)


class TemplatedSheet(metaclass=OrderedType):
    sheetname = Typed("sheetname", expected_type=str)
    active = Typed("active", expected_type=bool, value=False)
    workbook = Typed("workbook", expected_type=Workbook, allow_none=True)

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


DEFAULT_COLUMN_WIDTH = 8.43


class TableColumn:
    setter = Typed("setter", expected_type=FunctionType, allow_none=True)
    getter = Typed("getter", expected_type=FunctionType, allow_none=True)
    column_index = Typed("column_index", expected_type=int, allow_none=True)  # set by sheet

    # Rendering properties
    header = Typed("header", expected_type=str, allow_none=True)
    width = Typed("width", expected_types=(int, float), value=8.43)
    hidden = Typed("hidden", expected_type=bool, value=False)
    group = Typed("group", expected_type=bool, value=False)
    data_validation = Typed("data_validation", expected_type=DataValidation, allow_none=True)

    # Reading/writing properties
    default_value = None  # internal value
    allow_blank = Typed("allow_blank", expected_type=bool, value=True)

    BLANK_VALUES = (None, "")

    def __init__(self, object_attr=None, getter=None, setter=None, header=None, width=None, hidden=None, group=None,
                 data_validation=None, default_value=None, allow_blank=None):
        self.header = header if header is not None else self.header
        self.width = width if width is not None else self.width
        self.hidden = hidden if hidden is not None else self.hidden
        self.group = group if group is not None else self.group
        self.data_validation = data_validation if data_validation is not None else self.data_validation

        self.default_value = default_value if default_value is not None else self.default_value
        self.allow_blank = allow_blank if allow_blank is not None else self.allow_blank

        if object_attr:
            self.getter = lambda obj: getattr(obj, object_attr)
            self.setter = lambda obj, value: setattr(obj, object_attr, value)

        self.getter = getter if getter is not None else self.getter
        self.setter = setter if setter is not None else self.setter

    def get_value(self, obj):
        return self.getter(obj)

    def set_value(self, obj, value):
        self.setter(obj, value)

    def to_excel(self, value):
        return value

    def from_excel(self, cell):
        return cell.value

    def to_excel_with_blank_check(self, value):
        if value is None:
            if self.allow_blank:
                return None
            raise BlankNotAllowed()
        return self.to_excel(value)

    def from_excel_with_blank_check(self, cell):
        if cell.value in self.BLANK_VALUES:
            if not self.allow_blank:
                raise BlankNotAllowed(cell=cell)
            return self.default_value
        return self.from_excel(cell)

    def style_worksheet(self, worksheet, column_dimension):
        if self.width is not None:
            column_dimension.width = self.width

        column_dimension.hidden = self.hidden

        if self.data_validation:
            worksheet.add_data_validation(self.data_validation)

    def create_cell(self, worksheet, obj=None):
        return WriteOnlyCell(
            worksheet,
            value=self.to_excel_with_blank_check(
                self.get_value(obj) if obj is not None else self.default_value
            )
        )


class TableSheet(TemplatedSheet):
    item_class = TableColumn

    freeze_header = Typed("freeze_header", expected_type=bool, value=True)
    hide_excess_columns = Typed("hide_excess_columns", expected_type=bool, value=True)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.columns = list(self._items.values())

    def write(self, title=None, description=None, objects=None):
        pass

    def read(self, exception_policy):
        pass


class TemplatedWorkbook(Workbook, metaclass=OrderedType):
    item_class = TemplatedSheet

    templated_sheets = None
    template_styles = Typed("template_styles", expected_type=StyleSet)

    def __new__(cls, *args, file=None, **kwargs):
        if file:
            return load_workbook(file)
        return super().__new__(cls)

    def __init__(self, template_styles=None):
        super().__init__()

        self.templated_sheets = list(self._items.values())
        self.template_styles = template_styles or self.template_styles or StandardStyleSet()

    def remove_ordinary_sheets(self):
        templated_sheet_names = {sheet.sheetnamn for sheet in self.templated_sheets}
        for sheetname in templated_sheet_names:
            del self[sheetname]
