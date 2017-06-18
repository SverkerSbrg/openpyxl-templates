from collections import Counter
from itertools import chain
from types import FunctionType

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl_templates.exceptions import BlankNotAllowed, OpenpyxlTemplateException
from openpyxl_templates.style import StyleSet, StandardStyleSet
from openpyxl_templates.utils import Typed, OrderedType



