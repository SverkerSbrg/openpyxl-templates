from collections import deque
from itertools import chain

from openpyxl.styles import NamedStyle

from openpyxl_templates.utils import SolidFill


DEFAULT_ACCENT_COLOR = "1a1f43"


class ParentForExtendedStyleNotFound(KeyError):
    def __init__(self, extended_style):
        super().__init__("Base style '%s' for ExtendedStyle '%s' not found." % (extended_style.base, extended_style.name))


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

    def extend(self, parent):
        return NamedStyle(
            name=self.name,
            number_format=self.number_format or parent.number_format,
            fill=self.fill or parent.fill,
            font=self._extend_serializable(parent.font, self.font),
            border=self._extend_serializable(parent.border, self.border),
            alignment=self._extend_serializable(parent.alignment, self.alignment),
            protection=self._extend_serializable(parent.protection, self.protection),
        )

    @staticmethod
    def _extend_serializable(serializable, update):
        object_class = type(serializable)
        if object_class == type(update):
            return update
        kwargs = {}

        for attr in chain(object_class.__attrs__, object_class.__elements__):
            kwargs[attr] = getattr(serializable, attr)

        if update:
            kwargs.update({key: value for key, value in update.items() if value is not None})
        return object_class(**kwargs)


class StyleSet:
    _styles = None

    def __init__(self, *styles):
        self._styles = {}

        # Make names are only present once and that redeclared names takes precedence
        styles = {style.name: style for style in styles}.values()

        extended_styles = {}

        for style in styles:
            if isinstance(style, NamedStyle):
                self._add(style)
            elif isinstance(style, ExtendedStyle):
                extended_styles[style.name] = style
            else:
                raise TypeError("Unknown type")

        que = deque(extended_styles.values())

        while que:
            extended_style = que.pop()

            if extended_style.base in self:
                self._add(extended_style)
            elif extended_style.base in extended_styles:
                que.appendleft(extended_style)
            else:
                raise ParentForExtendedStyleNotFound(extended_style)

    def __getitem__(self, item):
        return self._styles[item]

    def __contains__(self, key):
        return key in self._styles

    def _add(self, style):
        if issubclass(type(style), ExtendedStyle):
            if style.base not in self:
                raise ParentForExtendedStyleNotFound(style)

            style = style.extend(self[style.base])

        if not isinstance(style, NamedStyle):
            raise ValueError("StyleSet can only handle NamedStyles")

        if style.name in self:
            raise ValueError("Style already exists")

        self._styles[style.name] = style

    @property
    def names(self):
        return tuple(style.name for style in self._styles.values())


class DefaultStyleSet(StyleSet):
    def __init__(self, *styles):
        super().__init__(
            NamedStyle(
                name="Default",
            ),
            ExtendedStyle(
                base="Default",
                name="Empty",
            ),
            ExtendedStyle(
                base="Empty",
                name="Title",
                font={"size": 20}
            ),
            ExtendedStyle(
                base="Empty",
                name="Description",
                font={"color": "FF777777"},
                alignment={"wrap_text": True}
            ),
            ExtendedStyle(
                base="Default",
                name="Header",
                font={"bold": True, "color": "FFFFFFFF"}, fill=SolidFill(DEFAULT_ACCENT_COLOR)
            ),
            ExtendedStyle(
                base="Header",
                name="Header, center",
                alignment={"horizontal": "center"}
            ),
            ExtendedStyle(
                base="Default",
                name="Row"
            ),
            ExtendedStyle(
                base="Row",
                name="Row, string",
                number_format="@",
            ),
            ExtendedStyle(
                base="Row",
                name="Row, text",
                number_format="@",
                alignment={"wrap_text": True}
            ),
            ExtendedStyle(
                base="Row",
                name="Row, integer",
                number_format="# ##0",
            ),
            ExtendedStyle(
                base="Row",
                name="Row, decimal",
                number_format="0.00",
            ),
            ExtendedStyle(
                base="Row",
                name="Row, date",
                alignment={"horizontal": "center"},
                number_format="yyyy-mm-dd"
            ),
            ExtendedStyle(
                base="Row",
                name="Row, time",
                alignment={"horizontal": "center"},
                number_format="h:mm"
            ),
            *styles
        )
