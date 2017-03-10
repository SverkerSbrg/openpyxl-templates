from collections import OrderedDict
from itertools import chain

from openpyxl.styles import NamedStyle

from openpyxl_templates.utils import SolidFill, Typed


class StylesMetaClass(type):
    @classmethod
    def __prepare__(mcs, name, bases):
        return OrderedDict()

    def __new__(mcs, name, bases, classdict):
        result = type.__new__(mcs, name, bases, dict(classdict))

        result._named_styles = OrderedDict(
            [
                (attr, column)
                for attr, column
                in classdict.items()
                if column.__class__ in (NamedStyle, ExtendedStyle)
            ]
        )

        return result


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


class StyleSet(dict, metaclass=StylesMetaClass):
    default = Typed("default", expected_type=NamedStyle, allow_none=True)

    def __init__(self, default=None, **styles):
        super().__init__()
        self.default = default

        self._names = set()
        self._styles = {}
        # self._style_hash_map = {}

        for key, style in chain(self._named_styles.items(), styles.items()):
            self[key] = style

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

            if style in self._styles:
                style = self._styles[style]
            else:
                self._styles[style] = style
                if style.name in self._names:
                    raise Exception("Duplicated styles names")
                else:
                    self._names.add(style.name)

        super().__setitem__(key, style)

    def __getitem__(self, key):
        if type(key) == str:
            try:
                return super().__getitem__(key)
            except KeyError:
                if self.default:
                    return self.default
                else:
                    raise Exception(
                        "The key '%s' is not associated with a style. "
                        "Try one of the existing keys or assign a default style. Existing keys: %s."
                        % (key, str(list(key for key in self.keys())))
                    )
        elif key is None:
            return None

        raise TypeError("'%s' is an invalid style key. It must be either a string or None." % key)

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


class StandardStyleSet(StyleSet):
    default = NamedStyle(name="Default")
    header = ExtendedStyle(base="default", name="HeaderX", font={"bold": True, "color": "FFFFFFFF"}, fill=SolidFill("5d1738"))
    header_center = ExtendedStyle(base="header", name="Header, center", alignment={"horizontal": "center"})
    row = default
    row_char = ExtendedStyle(
        base="row",
        name="Row, string",
        number_format="@",
    )
    row_text = ExtendedStyle(
        base="row",
        name="Row, text",
        number_format="@",
        alignment={"wrap_text": True}
    )
    row_integer = ExtendedStyle(
        base="row",
        name="Row, integer",
        number_format="# ##0",
    )
    row_float = ExtendedStyle(
        base="row",
        name="Row, decimal",
        number_format="0.00",
    )
    row_date = ExtendedStyle(
        base="row",
        name="Row, date",
        alignment={"horizontal": "center"},
        number_format="yyyy-mm-dd"
    )
    row_time = ExtendedStyle(
        base="row",
        name="Row, time",
        alignment={"horizontal": "center"},
        number_format="h:mm"
    )


class ColumnStyleMixin():
    header_style = Typed("header_style", expected_type=str, allow_none=True)
    row_style = Typed("row_style", expected_type=str, allow_none=True)

    def __init__(self, header_style=None, row_style=None):
        self.header_style = header_style or self.header_style
        self.row_style = row_style or self.row_style

    def inherit_styles(self, parent):
        self.header_style = self.header_style or parent.header_style
        self.row_style = self.row_style or parent.row_style


class SheetStyleMixin(ColumnStyleMixin):
    empty_style = Typed("empty_style", expected_type=str, allow_none=True)
    title_style = Typed("title_style", expected_type=str, allow_none=True)
    description_style = Typed("descriptor_style", expected_type=str, allow_none=True)

    def __init__(self, empty_style=None, title_style=None, description_style=None, **kwargs):
        self.empty_style = empty_style or self.empty_style
        self.title_style = title_style or self.title_style
        self.description_style = description_style or self.description_style

        super().__init__(**kwargs)

    def inherit_styles(self, parent):
        super().inherit_styles(parent)
        self.empty_style = self.empty_style or parent.empty_style
        self.title_style = self.title_style or parent.title_style
        self.description_style = self.description_style or parent.description_style

        if hasattr(self, "columns"):
            for column in self.columns:
                column.inherit_styles(self)

#
# def iter_font(font):
#     for attr in Font.__elements__:
#         value = getattr(font, attr)
#         if value:
#             yield attr, value
#
#
# def iter_alignment(alignment):
#     for attr in Alignment.__fields__:
#         value = getattr(alignment, attr)
#         if value not in (None, 0):
#             yield attr, value
#
#
# def iter_border(border):
#     for attr in border.__attrs__:
#         value = getattr(border, attr)
#         if value and attr != "outline":
#             yield attr, value
#         elif attr == "outline" and not value:
#             yield attr, value
#
#
# class CellStyle:
#     font = None
#     fill = None
#     alignment = None
#     border = None
#
#     def __init__(self, font=None, fill=None, alignment=None, border=None):
#         self.font = self._merge_fonts(self.font, font)
#         self.fill = self._merge_fills(self.fill, fill)
#         self.alignment = self._merge_alignments(self.alignment, alignment)
#         self.border = self._merge_borders(self.border, border)
#
#     def style_cell(self, cell):
#         cell.font = self.font
#         if self.fill is not None:
#             cell.fill = self.fill
#         cell.alignment = self.alignment
#         cell.border = self.border
#         return cell
#
#     @classmethod
#     def merge(cls, *styles):
#         return cls(
#             font=cls._merge_fonts(*(style.font for style in styles if style)),
#             fill=cls._merge_fills(*(style.fill for style in styles if style)),
#             alignment=cls._merge_alignments(*(style.alignment for style in styles if style)),
#             border=cls._merge_borders(*(style.border for style in styles if style))
#         )
#
#     @classmethod
#     def _merge_fonts(cls, *fonts):
#         result = Font()
#         for font in fonts:
#             if font is not None:
#                 for attr, value in iter_font(font):
#                     setattr(result, attr, value)
#         return result
#
#     @classmethod
#     def _merge_fills(cls, *fills):
#         for fill in fills[::-1]:
#             if fill is not None:
#                 return fill
#         return None
#
#     @classmethod
#     def _merge_alignments(cls, *alignments):
#         result = Alignment()
#         for aligment in alignments:
#             if aligment is not None:
#                 for attr, value in iter_alignment(aligment):
#                     setattr(result, attr, value)
#         return result
#
#     @classmethod
#     def _merge_borders(cls, *borders):
#         result = Border()
#         for border in borders:
#             if border is not None:
#                 for attr, value in iter_border(border):
#                     setattr(result, attr, value)
#         return result
#
#     def __str__(self):
#         return "Style: %s" % {
#             "font": {attr: value for attr, value in iter_font(self.font)},
#             "fill": str(self.fill),
#             "alignment": {attr: value for attr, value in iter_alignment(self.alignment)},
#             "border": {attr: value for attr, value in iter_border(self.border)}
#         }
#
#
# class ColumnStyle:
#     base = None
#     header = None
#     cell = None
#
#     _style_attrs = ("base", "cell", "header")
#
#     def __init__(self, base=None, header=None, cell=None):
#         self.base = CellStyle.merge(self.base, base)
#         self.cell = CellStyle.merge(self.base, self.cell, cell)
#         self.header = CellStyle.merge(self.base, self.header, header)
#
#     @classmethod
#     def merge(cls, *sheet_styles):
#         result = SheetStyle()
#         for attr in cls._style_attrs:
#             setattr(result, attr, CellStyle.merge(
#                 *[getattr(style, attr, None) for style in sheet_styles if style]
#             ))
#         return result
#
#     def style_cell(self, cell):
#         self.cell.style_cell(cell)
#
#     def style_header(self, cell):
#         self.header.style_cell(cell)
#
#
# class SheetStyle(ColumnStyle):
#     empty = None
#     title = None
#
#     _style_attrs = ColumnStyle._style_attrs + ("empty", "title")
#
#     def __init__(self, empty=None, title=None, **kwargs):
#         super().__init__(**kwargs)
#
#         self.empty = CellStyle.merge(self.base, self.empty, empty)
#         self.title = CellStyle.merge(self.base, self.title, title)
#
#     @property
#     def column_style(self):
#         return ColumnStyle(
#             base=self.base,
#             header=self.header,
#             cell=self.header
#         )
#
#     def style_empty(self, cell):
#         self.empty.style_cell(cell)
#
#     def style_title(self, cell):
#         self.title.style_cell(cell)
