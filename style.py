from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Font


def iter_font(font):
    for attr in Font.__elements__:
        value = getattr(font, attr)
        if value:
            yield attr, value


def iter_alignment(alignment):
    for attr in Alignment.__fields__:
        value = getattr(alignment, attr)
        if value not in (None, 0):
            yield attr, value


def iter_border(border):
    for attr in border.__attrs__:
        value = getattr(border, attr)
        if value and attr != "outline":
            yield attr, value
        elif attr == "outline" and not value:
            yield attr, value


class CellStyle:
    font = None
    fill = None
    alignment = None
    border = None

    def __init__(self, font=None, fill=None, alignment=None, border=None):
        self.font = self._merge_fonts(self.font, font)
        self.fill = self._merge_fills(self.fill, fill)
        self.alignment = self._merge_alignments(self.alignment, alignment)
        self.border = self._merge_borders(self.border, border)

    def style_cell(self, cell):
        cell.font = self.font
        if self.fill is not None:
            cell.fill = self.fill
        cell.alignment = self.alignment
        cell.border = self.border
        return cell

    @classmethod
    def merge(cls, *styles):
        return cls(
            font=cls._merge_fonts(*(style.font for style in styles if style)),
            fill=cls._merge_fills(*(style.fill for style in styles if style)),
            alignment=cls._merge_alignments(*(style.alignment for style in styles if style)),
            border=cls._merge_borders(*(style.border for style in styles if style))
        )

    @classmethod
    def _merge_fonts(cls, *fonts):
        result = Font()
        for font in fonts:
            if font is not None:
                for attr, value in iter_font(font):
                    setattr(result, attr, value)
        return result

    @classmethod
    def _merge_fills(cls, *fills):
        for fill in fills[::-1]:
            if fill is not None:
                return fill
        return None

    @classmethod
    def _merge_alignments(cls, *alignments):
        result = Alignment()
        for aligment in alignments:
            if aligment is not None:
                for attr, value in iter_alignment(aligment):
                    setattr(result, attr, value)
        return result

    @classmethod
    def _merge_borders(cls, *borders):
        result = Border()
        for border in borders:
            if border is not None:
                for attr, value in iter_border(border):
                    setattr(result, attr, value)
        return result

    def __str__(self):
        return "Style: %s" % {
            "font": {attr: value for attr, value in iter_font(self.font)},
            "fill": str(self.fill),
            "alignment": {attr: value for attr, value in iter_alignment(self.alignment)},
            "border": {attr: value for attr, value in iter_border(self.border)}
        }



