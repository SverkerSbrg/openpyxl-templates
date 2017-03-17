from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.styles.borders import BORDER_MEDIUM
from openpyxl.styles.fills import FILL_SOLID, PatternFill


def _color(color):
    if len(color) == 6:
        color = "FF%s" % color

    return color


def SolidFill(hex_color):
    fill = PatternFill(
        patternType=FILL_SOLID,
        fgColor=_color(hex_color)
    )
    return fill


def ColoredBorders(color, top=True, right=True, bottom=True, left=True):
    color = _color(color)
    return Border(
        top=Side(style=BORDER_MEDIUM, color=color) if top else None,
        right=Side(style=BORDER_MEDIUM, color=color) if right else None,
        bottom=Side(style=BORDER_MEDIUM, color=color) if bottom else None,
        left=Side(style=BORDER_MEDIUM, color=color) if left else None,
    )


# def NoBorders():
#     return Border(
#         top=None,
#         right=None,
#         # bottom=None,
#         left=None
#     )


class Typed(object):
    """Values must of a particular type"""
    name = None
    default_value = None
    expected_types = type(None)
    allow_none = False

    def __init__(self, name, value=None, expected_type=None, expected_types=None, allow_none=None):
        self.name = name

        if expected_types is not None:
            self.expected_types = expected_types
        else:
            self.expected_types = []

        if expected_type is not None:
            self.expected_types.append(expected_type)
        if allow_none is not None:
            self.allow_none = allow_none
        self.__doc__ = "Values must be of type {0}".format(self.expected_types)

        if value is not None:
            self.__set__(None, value)

    def __set__(self, instance, value):
        is_subclass = bool([True for t in self.expected_types if issubclass(type(value), t)])
        if not type(value) in self.expected_types and not is_subclass:
            if not self.allow_none or (self.allow_none and value is not None):
                raise TypeError("Attribute '%s' got type '%s' expected one of '%s'" % (
                self.name, type(value), str(self.expected_types)))

        if instance is not None:
            instance.__dict__[self.name] = value
        else:
            self.default_value = value

    def __get__(self, instance, owner):
        if instance is not None:
            try:
                return instance.__dict__[self.name]
            except KeyError:
                pass
        return self.default_value

    def __repr__(self):
        return self.__doc__
