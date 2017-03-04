from openpyxl.styles.fills import FILL_SOLID, PatternFill


def SolidFill(hex_color):
    fill = PatternFill(
        patternType=FILL_SOLID,
        fgColor="FF%s" % hex_color
    )
    return fill