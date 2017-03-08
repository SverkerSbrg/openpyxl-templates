from itertools import chain

from openpyxl.styles import NamedStyle

from openpyxl_templates.style import CellStyle


class WorkbookTemplate:
    sheets = None
    active_sheet = None

    styles = None

    # style = None
    # header_style = None
    # title_style = None
    # description_style = None

    def __init__(self, workbook, sheets=None, active_sheet=None, styles=None):
        self.workbook = workbook
        self.sheets = sheets or self.sheets or []
        self.active_sheet = active_sheet or self.active_sheet
        self.styles = styles or self.styles or []

        for sheet in self.sheets:
            if sheet.styles:
                for style in sheet.styles:
                    if style:
                        self.styles.append(style)

        self._style_set = StyleSet(*self.styles)

        self._sheet_map = {sheet.name: sheet for sheet in self.sheets}

    def write_sheet(self, name, objects):
        excel_sheet = self._sheet_map[name]
        worksheet = self.get_or_create_sheet(excel_sheet)
        excel_sheet.write(worksheet, self._style_set, objects)

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


class ExtendedStyle(dict):
    def __init__(self, base, name, font=None, fill=None, border=None, alignment=None, number_format=None,
                 protection=None):
        super().__init__()

        self.base = base
        self.name = name
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection


class StyleSet(dict):
    DEFAULT_STYLE = "__default__"
    DEFAULT_TITLE_STYLE = "__title__"
    DEFAULT_HEADER_STYLE = "__header__"
    DEFAULT_ROW_STYLE = "__row__"
    DEFAULT_STYLES = DEFAULT_STYLE, DEFAULT_TITLE_STYLE, DEFAULT_HEADER_STYLE, DEFAULT_ROW_STYLE

    def __init__(self, *styles, __default__=None):
        self.names = set()
        self.style_hash_map = {}
        super().__init__()

        if __default__:
            self[__default__.name] = __default__
            style_names = set((style.name for style in styles))
            for key in self.DEFAULT_STYLES:
                if key not in style_names:
                    self[key] = __default__

        for style in styles:
            self[style.name] = style

    def __setitem__(self, key, style):
        if type(key) != str:
            raise Exception("Keys must be strings")

        if key in self:
            raise Exception("Duplicate key")

        if style is not None:
            if type(style) == ExtendedStyle:
                style = self._extend_style(style)

            if type(style) != NamedStyle:
                raise ValueError("Must be named style")

            style_hash = self._hash_style_without_name(style)
            if style_hash in self.style_hash_map:
                style = self.style_hash_map[style_hash]
            else:
                self.style_hash_map[style_hash] = style

        super().__setitem__(key, style)

    def __getitem__(self, item):
        _type = type(item)
        if _type == str:
            try:
                return super().__getitem__(item)
            except KeyError:
                return None

        if _type not in (ExtendedStyle, NamedStyle):
            raise Exception("Unknown type")

        if item.name not in self:
            self[item.name] = item

        return super().__getitem__(item.name)

    def _extend_style(self, extended_style):
        base_style = self[extended_style.base]
        if not base_style:
            base_style = NamedStyle()

        return NamedStyle(
            name=extended_style.name,
            number_format=extended_style.number_format or base_style.number_format,
            fill=extended_style.fill or base_style.fill,
            font=self._extend_serializable(base_style.font, extended_style.font),
            border=self._extend_serializable(base_style.border, extended_style.border),
            alignment=self._extend_serializable(base_style.alignment, extended_style.alignment),
            protection=self._extend_serializable(base_style.protection, extended_style.protection),
        )

    def _extend_serializable(self, serializable, update):
        object_class = type(serializable)
        if object_class == type(update):
            return update
        kwargs = {}
        object_class = type(serializable)
        for attr in chain(object_class.__attrs__, object_class.__elements__):
            kwargs[attr] = getattr(serializable, attr)

        if update:
            kwargs.update({key: value for key, value in update.items() if value is not None})
        return object_class(**kwargs)

    def _hash_style_without_name(self, style):
        fields = []
        for attr in NamedStyle.__elements__ + ("number_format",):
            val = getattr(style, attr)
            if isinstance(val, list):
                val = tuple(val)
            fields.append(val)

        return hash(tuple(fields))
