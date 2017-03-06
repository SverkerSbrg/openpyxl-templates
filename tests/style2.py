from itertools import chain
from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Font

from openpyxl.styles import NamedStyle




class NamedStyleManager:

    @classmethod
    def merge(cls, *named_styles, result_name=None):
        values = {}

        named_styles = [style for style in named_styles if style]
        if len(named_styles) == 0:
            return None
        if len(named_styles) == 1:
            return named_styles[0]

        base_style = named_styles[0]

        for attr in NamedStyle.__elements__:
            values[attr] = cls.merge_serializable(*[getattr(style, attr) for style in named_styles])

        # result_name = result_name or
        result = NamedStyle(
            name=base_style.name,
            number_format=next((style.number_format for style in named_styles[::-1] if style.number_format), None),
            **values
        )
        if base_style == result:
            return base_style

        result.name = result_name
        return result


    @classmethod
    def merge_serializable(cls, *serializables):
        values = {}
        serializables = list([s for s in serializables if s])

        if not serializables:
            return None

        object_class = serializables[0].__class__
        for serializable in serializables:
            if object_class != serializable.__class__:
                raise TypeError("Cannot merge instances of different types")

            for attr in chain(object_class.__attrs__, object_class.__elements__):
                value = getattr(serializable, attr)
                if value not in (None,):
                    values[attr] = getattr(serializable, attr)

        return object_class(**values)


class WorksheetStyle:
    normal = NamedStyle(name="Normal")
    title = NamedStyle(name="Title")
    header = NamedStyle(name="Header")
    cell = NamedStyle(name="Cell")

    def __init__(self, normal=None, title=None, header=None, cell=None):
        self.normal = NamedStyleManager.merge(self.normal, normal, result_name="Normal")
        self.title = NamedStyleManager.merge(self.title, title, result_name="Title")
        self.header = NamedStyleManager.merge(self.header, header, result_name="Header")
        self.cell = NamedStyleManager.merge(self.cell, cell, result_name="Cell")




# class Style2(NamedStyle):
#
#     @classmethod
#     def combine(cls, first, second, result_name):
#
#
#         if not other
#         pass