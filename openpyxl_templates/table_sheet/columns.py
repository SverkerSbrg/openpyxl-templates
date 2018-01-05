from datetime import date, datetime, timedelta, time
from types import FunctionType

from collections import Iterable
from openpyxl.cell import WriteOnlyCell
from openpyxl.formatting import Rule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.exceptions import OpenpyxlTemplateException, CellException
from openpyxl_templates.utils import Typed, FakeCell


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
    conditional_formatting = Typed("conditional_formatting", expected_type=Rule, allow_none=True)

    # Reading/writing properties
    default = None  # internal value not excel
    allow_blank = Typed("allow_blank", expected_type=bool, value=True)
    ignore_forced_text = Typed("ignore_forced_text", expected_type=bool, value=True)

    header_style = Typed("header_style", expected_type=str, value="Header")
    row_style = Typed("row_style", expected_type=str, value="Row")
    freeze = Typed("freeze", expected_type=bool, value=False)

    BLANK_VALUES = (None, "")

    def __init__(self, header=None, object_attribute=None, source=None, width=None, hidden=None, group=None,
                 data_validation=None, conditional_formatting=None, default=None, allow_blank=None,
                 ignore_forced_text=None, header_style=None, row_style=None, freeze=False):

        self._header = header
        self.width = width
        self.hidden = hidden
        self.group = group
        self.data_validation = data_validation
        self.conditional_formatting = conditional_formatting

        self.default = default

        # Make sure the default value is valid
        if self.default is not None:
            self._to_excel(default)

        self.allow_blank = allow_blank
        self.ignore_forced_text = ignore_forced_text

        self._object_attribute = object_attribute
        self.source = source

        self.header_style = header_style
        self.row_style = row_style

        self.freeze = freeze

    def get_value_from_object(self, obj):
        if isinstance(obj, (list, tuple)):
            return obj[self.column_index - 1]

        if isinstance(obj, dict):
            return obj[self.object_attribute]

        return getattr(obj, self.object_attribute, None)

    def _to_excel(self, value):
        if value in self.BLANK_VALUES:
            if self.default is not None:
                return self.to_excel(self.default)
            if self.allow_blank:
                return None
            raise BlankNotAllowed(WriteOnlyCell())

        return self.to_excel(value)

    def to_excel(self, value):
        return value

    def _from_excel(self, cell):
        value = cell.value
        if self.ignore_forced_text and isinstance(value, str) and value.startswith("'"):
            value = value[1:]

        if value in self.BLANK_VALUES:
            if not self.allow_blank:
                raise BlankNotAllowed(cell=cell)
            return self.default

        return self.from_excel(cell, value)

    def from_excel(self, cell, value):
        return value

    def prepare_worksheet(self, worksheet):
        if self.data_validation:
            worksheet.add_data_validation(self.data_validation)

    def create_header(self, worksheet):
        header = WriteOnlyCell(ws=worksheet, value=self.header)
        if self.header_style:
            header.style = self.header_style
        return header

    def create_cell(self, worksheet, value=None):
        cell = WriteOnlyCell(
            worksheet,
            value=self._to_excel(value if value is not None else self.default)
        )
        if self.row_style:
            cell.style = self.row_style
        return cell

    def post_process_cell(self, worksheet, cell):
        pass

    def post_process_worksheet(self, worksheet, first_row, last_row, data_range):
        column_dimension = worksheet.column_dimensions[self.column_letter]

        # Hiding of grouped columns is handled on worksheet level.
        if not self.group:
            column_dimension.hidden = self.hidden
        column_dimension.width = self.width

        if self.data_validation:
            self.data_validation.ranges.append(data_range)

        if self.conditional_formatting:
            worksheet.conditional_formatting.add(data_range, self.conditional_formatting)

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

    def __str__(self):
        return "%s(%s)" % (self.__class__.__name__, self._header or self._object_attribute or "")

    def __repr__(self):
        return str(self)


class StringToLong(CellException):
    def __init__(self, cell):
        super().__init__(
            "Value '%s' in cell '%s' is too long." % (cell.value, cell.coordinate)
        )


class CharColumn(TableColumn):
    max_length = Typed("max_length", expected_type=int, allow_none=True)

    def __init__(self, header=None, max_length=None, **kwargs):
        super().__init__(header=header, **kwargs)

        self.max_length = max_length

    def from_excel(self, cell, value):
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
    def __init__(self, **kwargs):
        kwargs.setdefault("row_style", "Row, text")
        super().__init__(**kwargs)


class UnableToParseException(CellException):
    type = None

    def __init__(self, cell=None, value=None):
        if cell:
            message = "Unable to convert value '%s' of cell '%s' to %s." % (cell.value, cell.coordinate, self.type)
        else:
            message = "Unable to convert value '%s' to '%s'" % (value, self.type)
        super().__init__(
            message
        )


class UnableToParseBool(UnableToParseException):
    type = "boolean"


class BoolColumn(TableColumn):
    excel_true = Typed(name="excel_true", value=True, expected_types=(str, int, float, bool))
    excel_false = Typed(name="excel_false", value=False, expected_types=(str, int, float, bool))

    list_validation = Typed("list_validation", expected_type=bool, value=True)
    strict = Typed("strict", expected_type=bool, value=False)

    def __init__(self, header=None,  excel_true=None, excel_false=None, list_validation=None, strict=None, **kwargs):
        self.excel_true = excel_true
        self.excel_false = excel_false
        self.list_validation = list_validation
        self.strict = strict

        super().__init__(header=header, **kwargs)

        if self.list_validation and not self.data_validation:
            self.data_validation = DataValidation(
                type="list",
                formula1="\"%s\"" % ",".join((
                    str(self.excel_true if self.excel_true is not True else "TRUE"),
                    str(self.excel_false if self.excel_false is not False else "FALSE")
                ))
            )

    def to_excel(self, value):
        return self.excel_true if value else self.excel_false

    def from_excel(self, cell, value):
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
    def __init__(self, **kwargs):
        kwargs.setdefault("row_style", "Row, decimal")
        kwargs.setdefault("default", 0.0)
        super().__init__(**kwargs)

    def to_excel(self, value):
        try:
            return float(value)
        except (ValueError, TypeError):
            raise UnableToParseFloat(value=value)

    def from_excel(self, cell, value):
        try:
            return float(value)
        except (ValueError, TypeError):
            raise UnableToParseFloat(cell=cell)


class UnableToParseInt(UnableToParseException):
    type = "int"


class RoundingRequired(CellException):
    def __init__(self, cell):
        super().__init__(
            "The value '%s'  in cell '%s' cannot be converted to an integer without rounding the value. Enable "
            "round_value to do this automatically." % (cell.value, cell.coordinate)
        )


class IntColumn(FloatColumn):
    round_value = Typed("round_value", expected_type=bool, value=True)

    def __init__(self, header=None, round_value=None, **kwargs):
        kwargs.setdefault("row_style", "Row, integer")
        kwargs.setdefault("default", 0)
        super().__init__(header=header, **kwargs)

        self.round_value = round_value

    def to_excel(self, value):
        try:
            f = float(value)
            i = round(f, 0)
            if i != f and not self.round_value:
                raise RoundingRequired(FakeCell(value=value))
            return int(i)
        except (ValueError, TypeError):
            raise UnableToParseInt(value=value)

    def from_excel(self, cell, value):
        try:
            f = float(value)
            i = round(f, 0)
            if i != f and not self.round_value:
                raise RoundingRequired(cell=cell)
            return int(i)
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
    list_validation = Typed(name="list_validation", value=True, expected_type=bool)

    choices = Typed(name="choices", expected_type=Iterable)
    to_excel_map = None
    from_excel_map = None

    def __init__(self, header=None, choices=None, list_validation=None, **kwargs):

        self.choices = tuple(choices) if choices else None
        self.list_validation = list_validation

        self.to_excel_map = {internal: excel for internal, excel in self.choices}
        self.from_excel_map = {excel: internal for internal, excel in self.choices}

        # Setup maps before super().__init__() to validation of default value.
        super().__init__(header=header, **kwargs)

        if self.list_validation and not self.data_validation:
            self.data_validation = DataValidation(
                type="list",
                formula1="\"%s\"" % ",".join('%s' % str(excel) for internal, excel in self.choices)
            )

    def to_excel(self, value):
        if value not in self.to_excel_map:
            if self.default is not None:
                value = self.default

            if value not in self.to_excel_map:
                raise IllegalChoice(FakeCell(value), tuple(self.to_excel_map.keys()))

        return self.to_excel_map[value]

    def from_excel(self, cell, value):
        if value not in self.from_excel_map:
            if self.default is not None:
                return self.default

            if value not in self.from_excel_map:
                raise IllegalChoice(cell, tuple(self.from_excel_map.keys()))

        return self.from_excel_map[value]


class FortnumChoiceColumn(ChoiceColumn):
    def __init__(self, fortnum, **kwargs):
        kwargs["choices"] = ((f, str(f)) for f in fortnum)
        super().__init__(**kwargs)


class UnableToParseDatetime(UnableToParseException):
    type = "datetime"


class DatetimeColumn(TableColumn):
    SECONDS_PER_DAY = 24 * 60 * 60

    def __init__(self, **kwargs):
        kwargs.setdefault("row_style", "Row, date")
        kwargs.setdefault("header_style", "Header, center")
        super().__init__(**kwargs)

    def from_excel(self, cell, value):
        if isinstance(value, (datetime, date)):
            return value

        if type(value) in (int, float):
            # Excel dates start at 1900-01-01
            if value < 1:
                raise UnableToParseDatetime(cell)

            result = datetime(year=1900, month=1, day=1) + timedelta(days=value - 2)

            # Excel incorrectly assumes 1900 to be a leap year.
            if value < 61:
                result += timedelta(days=1)
            return result

        raise UnableToParseDatetime(cell)

    def to_excel(self, value):
        if type(value) == date:
            value = datetime.combine(value, time.min)
        if not isinstance(value, datetime):
            raise UnableToParseDatetime(value=value)

        delta = (value - datetime(year=1900, month=1, day=1, tzinfo=value.tzinfo))
        value = delta.days + delta.seconds / self.SECONDS_PER_DAY + 2

        # Excel incorrectly assumes 1900 to be a leap year.
        if value < 61:
            if value < 1:
                raise UnableToParseDatetime(value=value)
            value -= 1
        return value


class UnableToParseDate(UnableToParseException):
    type = "Row, date"


class DateColumn(DatetimeColumn):
    def from_excel(self, cell, value):
        try:
            return super().from_excel(cell, value).date()
        except UnableToParseDatetime:
            raise UnableToParseDate(cell=cell)

    def to_excel(self, value):
        return int(super().to_excel(value))


class UnableToParseTime(UnableToParseException):
    type = "time"


class TimeColumn(DatetimeColumn):
    def __init__(self, **kwargs):
        kwargs.setdefault("row_style", "Row, time")
        super().__init__(**kwargs)

    def from_excel(self, cell, value):
        if type(value) == time:
            return value

        try:
            return super().from_excel(cell, value).time()
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
    formula = Typed(name="formula", expected_type=str, allow_none=True)

    def __init__(self, formula=None, **kwargs):
        self.formula = formula

        if not self.formula:
            raise NoFormula()

        super().__init__(**kwargs)

    def get_value_from_object(self, obj):
        return self.formula


class EmptyColumn(TableColumn):
    def get_value_from_object(self, obj):
        return None
