from types import FunctionType

from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.exceptions import OpenpyxlTemplateException, BlankNotAllowed
from openpyxl_templates.utils import Typed


class ColumnIndexNotSet(OpenpyxlTemplateException):
    def __init__(self, column):
        super().__init__(
            "Column index not set for column '%s'. This should be done automatically by the TableSheet." % column
        )


class ObjectAttributeNotSet(OpenpyxlTemplateException):
    def __init__(self, column):
        super().__init__(
            "object_attribute not set for column '%s'. This should be done automatically by the TableSheet. "
            "The attributed must be assigned explicitly if added after class declaration" % column
        )

DEFAULT_COLUMN_WIDTH = 8.43


class TableColumn:
    _object_attribute = Typed("_object_attribute", expected_type=str, allow_none=True)
    source = Typed("source", expected_types=(str, FunctionType), allow_none=True)
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

    header_style = Typed("header_style", expected_type=str, value="Header")
    row_style = Typed("row_style", expected_type=str, value="Row")

    BLANK_VALUES = (None, "")

    def __init__(self, object_attribute=None, source=None, header=None, width=None, hidden=None, group=None,
                 data_validation=None, default_value=None, allow_blank=None):
        self._header = header if header is not None else self._header
        self.width = width if width is not None else self.width
        self.hidden = hidden if hidden is not None else self.hidden
        self.group = group if group is not None else self.group
        self.data_validation = data_validation if data_validation is not None else self.data_validation

        self.default_value = default_value if default_value is not None else self.default_value
        self.allow_blank = allow_blank if allow_blank is not None else self.allow_blank

        self._object_attribute = object_attribute if object_attribute is not None else self._object_attribute
        self.source = source if source is not None else self.source

    # def get_value(self, obj):
    #     return self.getter(obj)
    #
    # def set_value(self, obj, value):
    #     self.setter(obj, value)

    def get_value_from_object(self, object):
        if isinstance(object, (list, tuple)):
            return object[self.column_index]

        if isinstance(object, dict):
            return object[self.object_attribute]

        return getattr(object, self.object_attribute, None)

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
        if self.data_validation:
            worksheet.add_data_validation(self.data_validation)

    def create_header(self, worksheet):
        header = WriteOnlyCell(ws=worksheet, value=self.header)
        header.style = self.header_style
        return header

    def create_cell(self, worksheet, value=None):
        cell = WriteOnlyCell(
            worksheet,
            value=self.to_excel_with_blank_check(value or self.default_value)
        )
        cell.style = self.row_style
        return cell

    def post_process_cell(self, worksheet, cell):
        if self.data_validation:
            self.data_validation.add(cell)

    def post_process_worksheet(self, worksheet):
        column_dimension = worksheet.column_dimensions[self.column_letter]
        column_dimension.hidden = self.hidden
        column_dimension.width = self.width

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

    @property
    def object_attribute(self):
        if self._object_attribute is None:
            raise ObjectAttributeNotSet(self)

        return self._object_attribute



# class FormulaColumn(Column):
