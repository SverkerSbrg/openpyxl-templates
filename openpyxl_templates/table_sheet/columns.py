from datetime import date, datetime, timedelta, time
from types import FunctionType

from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.exceptions import OpenpyxlTemplateException, CellException
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


class BlankNotAllowed(CellException):
    def __init__(self, cell):
        super().__init__("The cell '%s' is not allowed to be empty." % cell.coordinate)


class TableColumn:
    _object_attribute = Typed("_object_attribute", expected_type=str, allow_none=True)
    source = Typed("source", expected_types=(str, FunctionType), allow_none=True)
    _column_index = None

    # Rendering properties
    _header = Typed("header", expected_type=str, allow_none=True)
    width = Typed("width", expected_types=(int, float), value=DEFAULT_COLUMN_WIDTH * 2)
    hidden = Typed("hidden", expected_type=bool, value=False)
    group = Typed("group", expected_type=bool, value=False)
    data_validation = Typed("data_validation", expected_type=DataValidation, allow_none=True)

    # Reading/writing properties
    default = None  # internal value not excel
    allow_blank = Typed("allow_blank", expected_type=bool, value=True)

    header_style = Typed("header_style", expected_type=str, value="Header")
    row_style = Typed("row_style", expected_type=str, value="Row")

    BLANK_VALUES = (None, "")

    def __init__(self, object_attribute=None, source=None, header=None, width=None, hidden=None, group=None,
                 data_validation=None, default_value=None, allow_blank=None, header_style=None, row_style=None):
        self._header = header if header is not None else self._header
        self.width = width if width is not None else self.width
        self.hidden = hidden if hidden is not None else self.hidden
        self.group = group if group is not None else self.group
        self.data_validation = data_validation if data_validation is not None else self.data_validation

        self.default = default_value if default_value is not None else self.default
        self.allow_blank = allow_blank if allow_blank is not None else self.allow_blank

        self._object_attribute = object_attribute if object_attribute is not None else self._object_attribute
        self.source = source if source is not None else self.source

        self.header_style = header_style if header_style is not None else self.header_style
        self.row_style = row_style if row_style is not None else self.row_style

    def get_value_from_object(self, object):
        if isinstance(object, (list, tuple)):
            return object[self.column_index-1]

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
            return self.default
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
            value=self.to_excel_with_blank_check(value or self.default)
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
        return self._header or self._object_attribute or "Column%d" % self.column_index

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


class StringToLong(CellException):
    def __init__(self, cell):
        super().__init__(
            "Value '%s' in cell '%s' is too long." % (cell.value, cell.coordinate)
        )


# TODO: Add ability to force text, Eg. append and strip "'"
class CharColumn(TableColumn):
    max_length = Typed("max_length", expected_type=int, allow_none=True)

    def __init__(self, max_length=None, **kwargs):
        super().__init__(**kwargs)

        self.max_length = max_length if max_length is not None else self.max_length

    def from_excel(self, cell):
        value = cell.value
        if value is None:
            return None

        value = str(value)

        if self.max_length is not None and len(value) > self.max_length:
            raise StringToLong(cell)

        return value

    def to_excel(self, value):
        if value is None:
            return ""

        return str(value)


class TextColumn(CharColumn):
    row_style = "Row, text"


class UnableToParseException(CellException):
    type = None

    def __init__(self, cell):
        super().__init__(
            "Unable to convert value '%s' of cell '%s' to %s." % (cell.value, cell.coordinate, self.type)
        )


class UnableToParseBool(UnableToParseException):
    type = "boolean"


class BoolColumn(TableColumn):
    excel_true = "TRUE"
    excel_false = "FALSE"

    list_validation = Typed("list_validation", expected_type=bool, value=True)
    strict = Typed("strict", expected_type=bool, value=False)

    def __init__(self, excel_true=None, excel_false=None, list_validation=None, strict=None, **kwargs):
        self.excel_true = excel_true if excel_true is not None else self.excel_true
        self.excel_false = excel_false if excel_false is not None else self.excel_false
        self.list_validation = list_validation if list_validation is not None else self.list_validation
        self.strict = strict if strict is not None else self.strict

        if self.list_validation and not self.data_validation:
            self.data_validation = DataValidation(
                type="list",
                formula1="\"%s\"" % ",".join((self.excel_true, self.excel_false))
            )

        super().__init__(**kwargs)

    def to_excel(self, value):
        return self.excel_true if value else self.excel_false

    def from_excel(self, cell):
        value = cell.value

        if isinstance(value, bool):
            return value

        if value == self.excel_true:
            return True

        if value == self.excel_false:
            return False

        if self.strict:
            raise UnableToParseBool(cell)

        return bool(value)


class UnableToParseFloat(UnableToParseException):
    type = "float"


class FloatColumn(TableColumn):
    row_style = "Row, decimal"
    default_value = 0.0

    def to_excel(self, value):
        return float(value)

    def from_excel(self, cell):
        value = cell.value

        try:
            return float(value)
        except (ValueError, TypeError):
            raise UnableToParseFloat(cell)


class UnableToParseInt(UnableToParseException):
    type = "int"


class RoundingRequired(CellException):
    def __init__(self, cell):
        super().__init__(
            "The value '%s'  in cell '%s' cannot be converted to an integer without rounding the value. Enable "
            "round_value to do this automatically." % (cell.value, cell.coordinate)
        )


class IntColumn(FloatColumn):
    row_style = "Row, integer"
    default_value = 0
    round_value = Typed("round_value", expected_type=bool, value=True)

    def __init__(self, round_value=None, **kwargs):
        self.round_value = round_value if round_value is not None else self.round_value
        super().__init__(**kwargs)

    def to_excel(self, value):
        return int(value)

    def from_excel(self, cell):
        value = cell.value
        try:
            f = float(value)
            i = int(value)
            if i != f and not self.round_value:
                raise RoundingRequired(cell)
        except (ValueError, TypeError):
            raise UnableToParseInt(cell)


class IllegalChoice(CellException):
    def __init__(self, cell, choices):
        super().__init__(
            "The value '%s' in cell '%s' is not a legal choices. Choices are %s." % (
                cell.value,
                cell.coordinate,
                choices
            )
        )


class ChoiceColumn(TableColumn):
    choices = None  # ((internal_value, excel_value),)
    list_validation = True

    to_excel_map = None
    from_excel_map = None

    def __init__(self, choices=None, list_validation=None, **kwargs):
        super().__init__(**kwargs)

        self.choices = choices if choices is not None else self.choices
        self.list_validation = list_validation if list_validation is not None else self.list_validation

        self.to_excel_map = {internal: excel for excel, internal in self.choices}
        self.from_excel_map = {excel: internal for excel, internal in self.choices}

        if self.list_validation and not self.data_validation:
            self.data_validation = DataValidation(
                type="list",
                formula1="\"%s\"" % ",".join('%s' % str(excel) for excel, internal in self.choices)
            )

    def to_excel(self, value):
        return self.to_excel_map[value]

    def from_excel(self, cell):
        try:
            return self.from_excel_map[cell.value]
        except KeyError:
            raise IllegalChoice(cell, tuple(self.from_excel_map.keys()))


class UnableToParseDatetime(UnableToParseException):
    type = "datetime"


class DatetimeColumn(TableColumn):
    SECONDS_PER_DAY = 24 * 60 * 60

    row_style = "Row, date"
    header_style = "Header, center"

    def from_excel(self, cell):
        value = cell.value

        if isinstance(value, (datetime, date)):
            return value

        if type in (int, float):
            return datetime(year=1900, month=1, day=1) + timedelta(days=value - 2)

        raise UnableToParseDatetime(cell)

    def to_excel(self, value):
        if type(value) == date:
            value = datetime.combine(value, time.min)

        delta = (value - datetime(year=1900, month=1, day=1))
        return delta.days + delta.seconds / self.SECONDS_PER_DAY + 2


class UnableToParseDate(UnableToParseException):
    type = "Row, date"


class DateColumn(DatetimeColumn):
    def from_excel(self, cell):
        try:
            return super().from_excel(cell).date()
        except UnableToParseDatetime:
            raise UnableToParseDate(cell=cell)

    def to_excel(self, value):
        return int(super().to_excel(value))


class UnableToParseTime(UnableToParseException):
    type = "time"


class TimeColumn(DatetimeColumn):
    row_style = "Row, time"

    def from_excel(self, cell):
        if type(cell.value) == time:
            return cell.value

        try:
            return super().from_excel(cell).time()
        except UnableToParseDatetime:
            raise UnableToParseTime(cell)

    def to_excel(self, value):
        _type = type(value)

        if value is None:
            return None

        if _type == time:
            return value

        if _type == datetime:
            return value.time()

        if _type == date:
            return time.min


class NoFormula(OpenpyxlTemplateException):
    def __init__(self):
        super().__init__("No formula specified for FormulaColumn.")


class FormulaColumn(TableColumn):
    formula = None

    def __init__(self, formula=None, **kwargs):
        self.formula = formula if formula is not None else self.formula

        if not self.formula:
            raise NoFormula()

        super().__init__(**kwargs)

    def get_value_from_object(self, obj):
        return self.formula


class EmptyColumn(TableColumn):
    def get_value_from_object(self, obj):
        return None
