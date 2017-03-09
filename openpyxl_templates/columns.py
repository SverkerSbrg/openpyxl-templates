from datetime import date, timedelta, time
from datetime import datetime

from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.exceptions import BlankNotAllowed, IllegalMaxLength, MaxLengthExceeded, UnableToParseBool, \
    UnableToParseFloat, UnableToParseInt, IllegalChoice, UnableToParseDatetime, UnableToParseDate, UnableToParseTime
from openpyxl_templates.style import ColumnStyleMixin


class Column(ColumnStyleMixin):
    object_attr = None
    header = None
    width = None

    hidden = False
    data_validation = None
    default_value = None
    number_format = "General"
    allow_blank = True

    BLANK_VALUES = (None, "")

    def __init__(self, object_attr=None, header=None, width=None, hidden=None,
                 data_validation=None, default_value=None, number_format=None, allow_blank=None, **style_keys):
        super().__init__(**style_keys)

        self.object_attr = object_attr or self.object_attr
        self.header = header or self.header
        self.width = width or self.width

        if hidden is not None:
            self.hidden = hidden

        self.data_validation = data_validation or self.data_validation
        self.default_value = default_value or self.default_value
        self.number_format = number_format or self.number_format
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
                self.get_value_from_object(obj) if obj else self.default_value
            )
        )

    # def style_cell(self, cell, style_set):
    #     """Add styling, is separated from creation since data_validation must be applied after appending to worksheet"""
    #
    #     cell.number_format = self.number_format
    #
    #     if self.style:
    #         self.style.style_cell(cell)
    #
    #     if self.data_validation:
    #         self.data_validation.add(cell)
    #
    # def get_styled_header_cell(self, worksheet):
    #     return self.header_style.style_cell(WriteOnlyCell(worksheet, value=self.header))

    def style_worksheet(self, worksheet, column_dimension):
        if self.width is not None:
            column_dimension.width = self.width

        column_dimension.hidden = self.hidden

        if self.data_validation:
            worksheet.add_data_validation(self.data_validation)


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

        # if self.max_length is not None and len(value) > self.max_length:
        #     raise OpenpyxlTemplateCellException("String to long", None)

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
        return value

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

    def __init__(self, *args, choices=None, **kwargs):
        self.choices = choices or self.choices

        self.data_validation = DataValidation(
            type="list",
            formula1="\"%s\"" % ",".join('%s' % str(excel) for excel, internal in self.choices)
        )

        super().__init__(*args, **kwargs)

        self._to_excel_map = {internal: excel for excel, internal in choices}
        self._from_excel_map = {excel: internal for excel, internal in choices}

    def to_excel(self, value):
        return self._to_excel_map[value]

    def from_excel(self, cell):
        try:
            return self._from_excel_map[cell.value]
        except:
            raise IllegalChoice(cell=cell)


class DateTimeColumn(Column):
    SECONDS_PER_DAY = 24 * 60 * 60

    class formats:
        SHORT_DATE = "yyyy-mm-dd"
        LONG_DATE = "DDDD, MMMM DD, ÅÅÅÅ"
        DATETIME = "yyyy-mm-dd h:mm:ss"
        TIME = "h:mm:ss"
        SHORT_TIME = "h:mm"

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
            # if type(value) not in (datetime, date):
            #     raise OpenpyxlTemplateCellException("Not datetime")
            # return value


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
