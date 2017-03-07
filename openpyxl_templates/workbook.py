from itertools import chain

from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import NamedStyle

from openpyxl_templates.style import CellStyle


class WorkbookTemplate:
    sheets = None
    active_sheet = None

    named_styles = None

    style=None
    header_style = None
    title_style = None
    description_style = None

    def __init__(self, workbook, sheets=None, active_sheet=None, style=None, header_style=None, title_style=None, description_style=None):
        self.workbook = workbook
        self.sheets = sheets or self.sheets or []
        self.active_sheet = active_sheet or self.active_sheet

        self.style = CellStyle.merge(self.style, style)
        self.header_style = CellStyle.merge(self.style, self.header_style, header_style)
        self.title_style = CellStyle.merge(self.style, self.title_style, title_style)
        self.description_style = CellStyle.merge(self.style, self.description_style, description_style)

        for sheet in self.sheets:
            sheet.style = CellStyle.merge(self.style, sheet.style)
            sheet.header_style = CellStyle.merge(self.header_style, sheet.header_style)
            sheet.title_style = CellStyle.merge(self.header_style, sheet.header_style)
            sheet.description_style = CellStyle.merge(self.description_style, sheet.description_style)
            sheet.rebase_column_styles()

        self._sheet_map = {sheet.name: sheet for sheet in self.sheets}

    def write_sheet(self, name, objects):
        excel_sheet = self._sheet_map[name]
        worksheet = self.get_or_create_sheet(excel_sheet)
        excel_sheet.write(worksheet, objects)

        self.update_active_sheet()

    def read_rows(self, name):
        excel_sheet = self._sheet_map[name]
        worksheet = self.get_or_create_sheet(excel_sheet)
        return excel_sheet.read_rows(worksheet)

    def get_or_create_sheet(self, excel_sheet):
        name = excel_sheet.name
        if name in self.workbook.sheetnames:
            return self.workbook[name]
        return self.workbook.create_sheet(excel_sheet.name)

    def update_active_sheet(self):
        for index, sheetname in enumerate(self.workbook.sheetnames):
            if sheetname == self.active_sheet:
                self.workbook.active = index
                return



class NamedStyleWorkbookTemplate(WorkbookTemplate):
    named_styles = None

    def __init__(self, named_styles=None, *args, **kwargs):
        super.__init__(*args, **kwargs)

        self.named_styles = named_styles or self.named_styles or [
            NamedStyle(name="Header", font=Font(bold=True)),
            NamedStyle(name="Title", font=Font(size=20, bold=True)),
            NamedStyle(name="Cell"),
            NamedStyle(name="Normalx"),
        ]

        self._named_style_map = {style.name: style for style in self.named_styles}


def serializable_to_dict(serializable):
    object_class = serializable.__class__
    values = {}
    for attr in chain(object_class.__attrs__, object_class.__elements__):
        value = getattr(serializable, attr)
        if value is not None:
            values[attr] = getattr(serializable, attr)
    return values



class InheritingFont(dict):
    def __init__(self, name=None, sz=None, b=None, i=None, charset=None,
                 u=None, strike=None, color=None, scheme=None, family=None, size=None,
                 bold=None, italic=None, strikethrough=None, underline=None,
                 vertAlign=None, outline=None, shadow=None, condense=None,
                 extend=None):
        super().__init__()

        self["name"] = name
        self["sz"] = sz
        self["b"] = b
        self["i"] = i
        self["charset"] = charset
        self["u"] = u
        self["strike"] = strike
        self["color"] = color
        self["scheme"] = scheme
        self["family"] = family
        self["size"] = size
        self["bold"] = bold
        self["italic"] = italic
        self["strikethrough"] = strikethrough
        self["underline"] = underline
        self["vertAlign"] = vertAlign
        self["outline"] = outline
        self["shadow"] = shadow
        self["condense"] = condense
        self["extend"] = extend

    def get_font(self, base_font):
        font_dict = serializable_to_dict(base_font)
        font_dict.update(self)
        return Font(**font_dict)


class InheritingAlignment(dict):
    def __init__(self, horizontal=None, vertical=None,
                 textRotation=0, wrapText=None, shrinkToFit=None, indent=0, relativeIndent=0,
                 justifyLastLine=None, readingOrder=0, text_rotation=None,
                 wrap_text=None, shrink_to_fit=None, mergeCell=None):
        super().__init__()

        self[" horizontal"] =  horizontal
        self[" vertical"] =  vertical
        self["textRotation"] = textRotation
        self["wrapText"] = wrapText
        self["shrinkToFit"] = shrinkToFit
        self["indent"] = indent
        self["relativeIndent"] = relativeIndent
        self["justifyLastLine"] = justifyLastLine
        self["readingOrder"] = readingOrder
        self["text_rotation"] = text_rotation
        self["wrap_text"] = wrap_text
        self["shrink_to_fit"] = shrink_to_fit
        self["mergeCell"] = mergeCell

    def get_alignment(self, base_alignment):
        font_dict = serializable_to_dict(base_alignment)
        font_dict.update(self)
        return Alignment(**font_dict)



class InheritingNamedStyle:
    def __init__(self, inheriting_from=None, name=None, font=None, alignment=None):
        self.name = name
        self.inheriting_from = inheriting_from

        self.font = font
        self.alignmnet = alignment

    # def get_named_style(self, workbook):
    #     base_style =


        # find base style, create new style, check if existing, add or use existing




