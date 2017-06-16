from collections import OrderedDict, Counter
from types import FunctionType

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.columns import Column
from openpyxl_templates.exceptions import BlankNotAllowed, OpenpyxlTemplateException
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


class ColumnIndexNotSet(OpenpyxlTemplateException):
    def __init__(self, column):
        super().__init__(
            "Column index not set for column '%s'. This should be done automatically by the TableSheet." % column
        )


DEFAULT_COLUMN_WIDTH = 8.43


class TableColumn:
    setter = Typed("setter", expected_type=FunctionType, allow_none=True)
    getter = Typed("getter", expected_type=FunctionType, allow_none=True)
    _column_index = None

    # Rendering properties
    _header = Typed("header", expected_type=str, allow_none=True)
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
        self._header = header if header is not None else self._header
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

    def prepare_worksheet(self, worksheet):
        #    AddDataValidations
        pass

    def create_header(self):
        pass

    def post_process_worksheet(self, worksheet):
        #     Hide
        #     SetWidth
        #     ColumnStyle
        pass

    def style_worksheet(self, worksheet, column_dimension): # TODO: Replace with post_process_worksheet
        if self.width is not None:
            column_dimension.width = self.width

        column_dimension.hidden = self.hidden

        # if self.data_validation:
        #     worksheet.add_data_validation(self.data_validation)

    def create_cell(self, worksheet, obj=None):
        return WriteOnlyCell(
            worksheet,
            value=self.to_excel_with_blank_check(
                self.get_value(obj) if obj is not None else self.default_value
            )
        )

    @property
    def header(self):
        return self._header or "Column%d" % self.column_index

    @property
    def column_index(self):
        if self._column_index is None:
            raise ColumnIndexNotSet(self)
        return self._column_index

    @column_index.setter
    def column_index(self, value):
        self._column_index = value

    @property
    def column_letter(self):
        return get_column_letter(self.column_index)


class ColumnHeadersNotUnique(OpenpyxlTemplateException):
    def __init__(self, columns):
        counter = Counter(column.header for column in columns)
        super().__init__("headers '%s' has been declared more then once in the same TableSheet" % tuple(
            header
            for (header, count)
            in counter.items()
            if count > 1
        ))


class TableSheet(TemplatedSheet):
    item_class = TableColumn

    freeze_header = Typed("freeze_header", expected_type=bool, value=True)
    hide_excess_columns = Typed("hide_excess_columns", expected_type=bool, value=True)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.columns = list(self._items.values())

        for index, column in enumerate(self.columns):
            column.column_index = index + 1  # Start as 1

        self._validate()

    def _validate(self):
        self._check_unique_column_headers()

    def _check_unique_column_headers(self):
        if len(set(column.header for column in self.columns)) < len(self.columns):
            raise ColumnHeadersNotUnique(self.columns)



    def write(self, title=None, description=None, objects=None):
        worksheet = self.worksheet

        self.prepare_worksheet(worksheet)
        self.write_title(worksheet, title)
        self.write_description(worksheet, description)
        self.write_headers(worksheet)
        self.write_rows(worksheet, objects)
        self.post_process_worksheet(worksheet)
        pass

    def prepare_worksheet(self, worksheet):
        # Columns.prepare_worksheet(worksheet)
        #    AddDataValidations
        pass

    def write_title(self, worksheet, title=None):
        pass

    def write_description(self, worksheet, description=None):
        pass

    def write_headers(self, worksheet):
        pass

    def write_rows(self, worksheet, objects=None):
        pass

    def post_process_worksheet(self, worksheet):
        # Columns.post_process_worksheet
        #     Hide
        #     SetWidth
        #     ColumnStyle
        # Group
        # FreezePane
        # FormatAsTable
        # HideExcessColumns
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
