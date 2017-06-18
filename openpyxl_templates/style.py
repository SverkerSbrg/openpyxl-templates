from collections import OrderedDict
from itertools import chain

from openpyxl.styles import NamedStyle

from openpyxl_templates.utils import SolidFill, Typed, ColoredBorders


class _Colors:
    DARK_RED = "5d1738"
    DARK_BLUE = "1a1f43"


class StylesMetaClass(type):
    @classmethod
    def __prepare__(mcs, name, bases):
        return OrderedDict()

    def __new__(mcs, name, bases, classdict):
        result = type.__new__(mcs, name, bases, dict(classdict))

        if not hasattr(result, "_named_styles"):
            result._named_styles = OrderedDict()

        for attr, column in classdict.items():
            if column.__class__ in (NamedStyle, ExtendedStyle):
                result._named_styles[attr] = column

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

    @property
    def names(self):
        return tuple(style.name for style in self.values())


class StandardStyleSet(StyleSet):
    default = NamedStyle(name="Default")
    empty = ExtendedStyle(
        base="Default",
        name="Empty",
        border=ColoredBorders("FFFFFFFF")
    )
    title = ExtendedStyle(
        base="Empty",
        name="Title",
        font={"size": 20},
    )
    description = ExtendedStyle(
        base="Empty",
        name="Description",
        font={"color": "FF777777"}
    )
    header = ExtendedStyle(
        base="Default",
        name="Header",
        font={"bold": True, "color": "FFFFFFFF"}, fill=SolidFill(_Colors.DARK_BLUE)
    )
    header_center = ExtendedStyle(
        base="Header",
        name="Header, center",
        alignment={"horizontal": "center"}
    )
    row = ExtendedStyle(
        base="Default",
        name="Row",
        protection={"locked": False}
    )
    row_char = ExtendedStyle(
        base="Row",
        name="Row, string",
        number_format="@",
    )
    row_text = ExtendedStyle(
        base="Row",
        name="Row, text",
        number_format="@",
        alignment={"wrap_text": True}
    )
    row_integer = ExtendedStyle(
        base="Row",
        name="Row, integer",
        number_format="# ##0",
    )
    row_float = ExtendedStyle(
        base="Row",
        name="Row, decimal",
        number_format="0.00",
    )
    row_date = ExtendedStyle(
        base="Row",
        name="Row, date",
        alignment={"horizontal": "center"},
        number_format="yyyy-mm-dd"
    )
    row_time = ExtendedStyle(
        base="Row",
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
