from datetime import date, timedelta, time
from datetime import datetime

from openpyxl.cell import WriteOnlyCell
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.exceptions import BlankNotAllowed, IllegalMaxLength, MaxLengthExceeded, UnableToParseBool, \
    UnableToParseFloat, UnableToParseInt, IllegalChoice, UnableToParseDatetime, UnableToParseDate, UnableToParseTime
from openpyxl_templates.style import ColumnStyleMixin
from openpyxl_templates.utils import Typed

DEFAULT_WIDTH = 8.43


class Column(ColumnStyleMixin):
    object_attr = Typed("object_attr", expected_type=str, allow_none=False)
    header = Typed("header", expected_type=str, allow_none=True)
    width = Typed("width", expected_types=(int, float), allow_none=True, value=DEFAULT_WIDTH)

    hidden = Typed("hidden", expected_type=bool, value=False)
    data_validation = Typed("data_validation", expected_type=DataValidation, allow_none=True)
    default_value = None
    allow_blank = Typed("allow_blank", expected_type=bool, value=True)

    BLANK_VALUES = (None, "")

    def __init__(self, object_attr=None, header=None, width=None, hidden=None,
                 data_validation=None, default_value=None, allow_blank=None, **style_keys):
        super().__init__(**style_keys)

        self.object_attr = object_attr or self.object_attr
        self.header = header or self.header
        self.width = width if width is not None else self.width

        if hidden is not None:
            self.hidden = hidden

        self.data_validation = data_validation or self.data_validation
        self.default_value = default_value or self.default_value
        self.allow_blank = allow_blank or self.allow_blank

    def get_value_from_object(self, obj):
        return getattr(obj, self.object_attr)

    def set_value_to_object(self, obj, value):
        setattr(obj, self.object_attr, value)

    def to_excel(self, value):
        raise NotImplementedError()

    def from_excel(self, cell):
        raise NotImplementedError()

    def from_excel_with_blank_check(self, cell):
        if cell.value in self.BLANK_VALUES:
            if not self.allow_blank:
                raise BlankNotAllowed(cell)
            return self.default_value
        return self.from_excel(cell)

    def create_cell(self, worksheet, obj=None):
        """Create a cell, object=None implies a request for an cell with default value"""
        return WriteOnlyCell(
            worksheet,
            value=self.to_excel(
                self.get_value_from_object(obj) if obj else None
            )
        )

    def style_worksheet(self, worksheet, column_dimension):
        if self.width is not None:
            column_dimension.width = self.width

        column_dimension.hidden = self.hidden

        if self.data_validation:
            worksheet.add_data_validation(self.data_validation)

    def __str__(self):

        return "%s: %s" % (
            self.__class__.__name__,
            "{%s}" % ", ".join(
                ": ".join(item) for item in (
                    ("object_attr", str(self.object_attr)),
                    ("header", str(self.header)),
                    ("header_style", str(self.header_style)),
                    ("row_style", str(self.row_style)),
                    ("allow_blank", str(self.allow_blank))
                )
            )
        )


class CharColumn(Column):
    max_length = None
    row_style = "row_text"

    def __init__(self, max_length=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if max_length is not None:
            if max_length <= 0:
                raise IllegalMaxLength(column=self, max_length=max_length)

            self.max_length = max_length

    def from_excel(self, cell):
        value = str(cell.value)
        if self.max_length is not None and len(value) > self.max_length:
            raise MaxLengthExceeded(cell=cell)

        return str(cell.value)

    def to_excel(self, value):
        if value is None:
            return ""

        value = str(value)

        return value


class TextColumn(CharColumn):
    row_style = "row_text"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)


class BooleanColumn(Column):
    default_value = False

    EXCEL_TRUE = "TRUE"
    EXCEL_FALSE = "FALSE"

    def to_excel(self, value):
        if value:
            return "TRUE"
        else:
            return "FALSE"

    def from_excel(self, cell):
        value = cell.value
        _type = type(value)

        if _type == bool:
            return value

        if value == self.EXCEL_TRUE:
            return True

        if value == self.EXCEL_FALSE:
            return False

        if _type in (int, float):
            return bool(value)

        try:
            return bool(float(value))
        except (ValueError, TypeError):
            pass

        raise UnableToParseBool(cell=cell)


class FloatColumn(Column):
    row_style = "row_float"
    default_value = 0.0

    def to_excel(self, value):
        return value

    def from_excel(self, cell):
        value = cell.value

        try:
            return float(value)
        except (ValueError, TypeError):
            raise UnableToParseFloat(cell=cell)


class IntegerColumn(Column):
    row_style = "row_integer"
    default_value = 0
    round_value = True

    def to_excel(self, value):
        return int(value)

    def from_excel(self, cell):
        value = cell.value
        try:
            f = float(value)
            i = round(f)
            if i != f and not self.round_value:
                raise ValueError()
            return i
        except (ValueError, TypeError):
            raise UnableToParseInt(cell=cell)


class ChoiceColumn(Column):
    choices = None
    add_list_validation = True

    def __init__(self, *args, choices=None, add_list_validation=None, **kwargs):
        self.choices = choices or self.choices
        self.add_list_validation = add_list_validation if add_list_validation is not None else self.add_list_validation

        if self.add_list_validation and not self.data_validation:
            self.data_validation = DataValidation(
                type="list",
                formula1="\"%s\"" % ",".join('%s' % str(excel) for excel, internal in self.choices)
            )

        super().__init__(*args, **kwargs)

        self._to_excel_map = {internal: excel for excel, internal in self.choices}
        self._from_excel_map = {excel: internal for excel, internal in self.choices}

    def to_excel(self, value):
        return self._to_excel_map[value]

    def from_excel(self, cell):
        try:
            return self._from_excel_map[cell.value]
        except:
            raise IllegalChoice(cell=cell)


class DateTimeColumn(Column):
    SECONDS_PER_DAY = 24 * 60 * 60

    row_style = "row_date"
    header_style = "header_center"

    def from_excel(self, cell):
        value = cell.value
        _type = type(value)

        if _type == datetime:
            return value

        if _type == date:
            return datetime(date)

        if _type in (int, float):
            return datetime(year=1900, month=1, day=1) + timedelta(days=value - 2)

        raise UnableToParseDatetime(cell=cell)

    def to_excel(self, value):
        try:
            if type(value) == date:
                value = datetime.combine(value, time.min)

            delta = (value - datetime(year=1900, month=1, day=1))
            return delta.days + delta.seconds / self.SECONDS_PER_DAY + 2
        except:
            pass


class DateColumn(DateTimeColumn):
    def from_excel(self, cell):
        try:
            return super().from_excel(cell).date()
        except UnableToParseDatetime:
            raise UnableToParseDate(cell=cell)

    def to_excel(self, value):
        return int(super().to_excel(value))


class TimeColumn(DateTimeColumn):
    row_style = "row_time"

    def from_excel(self, cell):
        if type(cell.value) == time:
            return cell.value

        try:
            return super().from_excel(cell).time()
        except UnableToParseTime:
            raise UnableToParseDate(cell=cell)

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
