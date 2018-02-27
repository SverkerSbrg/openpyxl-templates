from collections import OrderedDict
from weakref import WeakKeyDictionary

from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.styles.borders import BORDER_MEDIUM
from openpyxl.styles.fills import FILL_SOLID, PatternFill
from openpyxl.utils import column_index_from_string

MAX_COLUMN_INDEX = column_index_from_string("XFD")


def _color(color):
    if len(color) == 6:
        color = "FF%s" % color

    return color


def SolidFill(hex_color):
    color = _color(hex_color)
    fill = PatternFill(
        patternType=FILL_SOLID,
        fgColor=color,
        start_color=color,
        end_color=color
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


class Typed(object):
    name = None
    default_value = None
    _values = None

    def __init__(self, name, value=None, expected_type=None, expected_types=None, allow_none=False):
        self.name = name
        self._values = WeakKeyDictionary()

        if expected_types is not None:
            self.expected_types = expected_types
        else:
            self.expected_types = []

        if expected_type is not None:
            self.expected_types.append(expected_type)
        self.allow_none = allow_none
        self.__doc__ = "Values must be of type {0}".format(self.expected_types)

        if value is not None:
            self.validate(value)
            self.default_value = value

    def __set__(self, instance, value):
        if instance is None:
            return

        self.validate(value)

        if value is None:
            try:
                del self._values[instance]
            except KeyError:
                pass
            return

        self._values[instance] = value

    def validate(self, value):
        if value is None:
            if self.allow_none or self.default_value is not None:
                return
            raise ValueError("Attribute '%s' must not be None" % self.name)

        is_subclass = bool([True for t in self.expected_types if issubclass(type(value), t)])
        if not (type(value) in self.expected_types or is_subclass):
            raise TypeError("Attribute '%s' got type '%s' expected one of '%s'" % (
                self.name, type(value), str(self.expected_types)))

    def __get__(self, instance, owner):
        if instance is not None:
            try:
                return self._values[instance]
            except KeyError:
                pass
        return self.default_value

    def __repr__(self):
        return self.__doc__


class class_property(classmethod):
    def __get__(self, instance, owner):
        return super(class_property, self).__get__(instance, owner)()


class OrderedType(type):
    item_class = None
    _items = None

    @classmethod
    def __prepare__(mcs, name, bases):
        return OrderedDict()

    def __new__(mcs, name, bases, classdict):
        obj = super(OrderedType, mcs).__new__(mcs, name, bases, classdict)

        items = OrderedDict()

        for base in bases[::-1]:
            items.update(getattr(base, "_items", OrderedDict()))

        item_class = getattr(obj, "item_class", None)
        if item_class:
            for key, value in classdict.items():
                if issubclass(type(value), item_class):
                    items[key] = value

        obj._items = items

        if hasattr(obj, "__register_objects__"):
            obj.__register_objects__(obj, classdict)
        return obj


class FakeCell:
    coordinate = "A1"

    def __init__(self, value):
        self.value = value

    @classmethod
    def create(cls, values):
        return tuple(cls(value) for value in values)


def FakeCells(*values):
    return tuple(FakeCell(value) for value in values)