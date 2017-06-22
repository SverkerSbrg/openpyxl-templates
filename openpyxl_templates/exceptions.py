class OpenpyxlTemplateException(Exception):
    pass


class CellException(OpenpyxlTemplateException):
    pass


class RowException(OpenpyxlTemplateException):
    pass


class CellExceptions(RowException):
    def __init__(self, cell_exceptions):
        self.cell_exceptions = cell_exceptions
        super().__init__(
            "Failed to read row due to cell errors: %s" %
            ", ".join("\n    %s: '%s'" % (e.coordinate, str(e)) for e in self.cell_exceptions)
        )


class SheetException(OpenpyxlTemplateException):
    pass


class RowExceptions(SheetException):
    def __init__(self, row_exceptions):
        self.exceptions = row_exceptions
        super().__init__()
