from openpyxl.styles.fills import FILL_SOLID, PatternFill


def SolidFill(hex_color):
    fill = PatternFill(
        patternType=FILL_SOLID,
        fgColor="FF%s" % hex_color
    )
    return fill


class Typed(object):
    """Values must of a particular type"""
    value = None
    expected_types = type(None)
    allow_none = False

    def __init__(self,  expected_type=None, expected_types=None, allow_none=None):
        if expected_types is not None:
            self.expected_types = expected_types
        else:
            self.expected_types = []

        if expected_type is not None:
            self.expected_types.append(expected_type)
        if allow_none is not None:
            self.allow_none = allow_none
        self.__doc__ = "Values must be of type {0}".format(self.expected_types)

    def __set__(self, instance, value):
        is_subclass = bool([True for t in self.expected_types if issubclass(type(value), t)])
        if not type(value) in self.expected_types and not is_subclass:
            if not self.allow_none or (self.allow_none and value is not None):
                raise TypeError("Got type '%s' expected one of '%s'" % (type(value), str(self.expected_types)))
        self.value = value

    def __get__(self, instance, owner):
        return self.value

    def __repr__(self):
        return self.__doc__
