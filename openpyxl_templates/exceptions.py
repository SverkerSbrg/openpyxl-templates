class OpenpyxlTemplateException(Exception):
    message = None

    def __init__(self, message=None):
        super().__init__(self._get_message(message))

    def _get_message(self, message=None):
        assert (message or self.message)
        return (message or self.message).format(self=self)


class InvalidConfiguration(OpenpyxlTemplateException):
    pass


class CannotGroupLastVisibleColumnAndHideExcessColumns(InvalidConfiguration):
    message = "Grouping the last column will not render properly when excess columns are hidden."


class ColumnException(OpenpyxlTemplateException):
    def __init__(self, message=None, column=None):
        self.column = column
        super().__init__(message)


class ColumnBlankNotAllowed(ColumnException):
    pass


class IllegalMaxLength(ColumnException):
    message = "'{self.[max_length]}' is not valid it must be an integer larger than zero."

    def __init__(self, max_length, *args, **kwargs):
        self.max_length = max_length
        super().__init__(*args, **kwargs)


class OpenpyxlTemplateCellException(OpenpyxlTemplateException):
    def __init__(self, message=None, cell=None):
        self.cell = cell
        super().__init__(message)

    @property
    def coordinate(self):
        return self.cell.coordinate if self.cell else "CellUnknown"


class BlankNotAllowed(OpenpyxlTemplateCellException):
    message = "Cell({self[coordinate]}) is empty"


class MaxLengthExceeded(OpenpyxlTemplateCellException):
    message = "The value '{self.cell.value}' in cell({self.coordinate}) is too long'"


class UnableToParseBool(OpenpyxlTemplateCellException):
    message = "Unable to convert the value '{self.cell.value}' in cell({self.coordinate.}) to bool"


class UnableToParseFloat(OpenpyxlTemplateCellException):
    message = "Unable to convert the value '{self.cell.value}' in cell({self.coordinate.}) to float"


class UnableToParseInt(OpenpyxlTemplateCellException):
    message = "Unable to convert the value '{self.cell.value}' in cell({self.coordinate}) to int"


class IllegalChoice(OpenpyxlTemplateCellException):
    message = "The value '{self.cell.value}' in cell({self.coordinate}) is not a valid choice"


class UnableToParseDatetime(OpenpyxlTemplateCellException):
    message = "Unable to convert the value '{self.cell.value}' in cell({self.coordinate}) to datetime"


class UnableToParseDate(OpenpyxlTemplateCellException):
    message = "Unable to convert the value '{self.cell.value}' in cell({self.coordinate}) to date"


class UnableToParseTime(OpenpyxlTemplateCellException):
    message = "Unable to convert the value '{self.cell.value}' in cell({self.coordinate}) to time"


class OpenpyxlTemplateRowException(OpenpyxlTemplateException):
    def __init__(self, message):
        super().__init__(message)


class HeaderNotFound(OpenpyxlTemplateException):
    def __init__(self, sheetname, headers):
        super().__init__(
            "Header column not found on sheet '%s' either make sure that the following headers are "
            "present '%s' or set read_only_after_header to false." % (
                sheetname,
                ", ".join(headers)
            )
         )


class CellExceptions(OpenpyxlTemplateRowException):
    def __init__(self, cell_exceptions):
        self.cell_exceptions = cell_exceptions
        super().__init__(
            "Failed to read row due to cell errors: %s" %
            ", ".join("\n    %s: '%s'" % (e.coordinate, str(e)) for e in self.cell_exceptions)
        )
