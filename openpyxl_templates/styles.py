from itertools import chain

from openpyxl.styles import NamedStyle

from openpyxl_templates.utils import SolidFill


class _Colors:
    DARK_RED = "5d1738"
    DARK_BLUE = "1a1f43"


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

    def extend(self, named_style):
        return NamedStyle(
            name=self.name,
            number_format=self.number_format or named_style.number_format,
            fill=self.fill or named_style.fill,
            font=self._extend_serializable(named_style.font, self.font),
            border=self._extend_serializable(named_style.border, self.border),
            alignment=self._extend_serializable(named_style.alignment, self.alignment),
            protection=self._extend_serializable(named_style.protection, self.protection),
        )

    @staticmethod
    def _extend_serializable(serializable, update):
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


class StyleSet:
    _styles = None

    def __init__(self, *styles):
        self._styles = {}

        for style in styles:
            self.add(style)

    def __getitem__(self, item):
        return self._styles[item]

    def __contains__(self, key):
        return key in self._styles

    def add(self, style):
        if issubclass(type(style), ExtendedStyle):
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
    def __init__(self, accent_color=_Colors.DARK_BLUE):
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
                font={"bold": True, "color": "FFFFFFFF"}, fill=SolidFill(accent_color)
            ),
            ExtendedStyle(
                base="Header",
                name="Header, center",
                alignment={"horizontal": "center"}
            ),
            ExtendedStyle(
                base="Default",
                name="Row",
                protection={"locked": False}
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
            )
        )
