from openpyxl.styles import NamedStyle, Font

from openpyxl_templates.styles import StyleSet, ExtendedStyle, DefaultStyleSet
from openpyxl_templates.utils import SolidFill

demo_style = StyleSet(
    NamedStyle(
        name="Default",
        font=Font(
            name="Arial",
            size=12
        )
    ),
    NamedStyle(
        name="Header",
        font=Font(
            name="Arial",
            size=12,
            bold=True,
        ),
    )
)

demo_style = StyleSet(
    NamedStyle(
        name="Default",
        font=Font(
            name="Arial",
            size=12
        )
    ),
    ExtendedStyle(
        base="Default",  # Reference to the style defined above
        name="Header",
        font={
            "bold": True,
        }
    )
)

bad_example = StyleSet(
    NamedStyle(
        name="Default",
        font=Font(
            name="Arial",
            size=12
        )
    ),
    ExtendedStyle(
        base="Default",
        name="Header",
        font=Font(
            # Openpyxl will set name="Calibri" by default which will override name="Arial and break inheritance.
            bold=True
        )
    )
)

new_font_default_style_set = DefaultStyleSet(
    NamedStyle(  # Replace the existing "Default" font with a new one.
        name="Default",
        font=Font(
            name="Arial",
        )
    )
)