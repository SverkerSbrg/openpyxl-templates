# TODO: Test StyleSet2
from unittest import TestCase

from openpyxl.styles import NamedStyle, Font

from openpyxl_templates.style import ExtendedStyle


class ExtendedStyleTests(TestCase):
    def setUp(self):
        self.style = NamedStyle(name="base", font=Font(bold=True, size=12))

    def test_extend(self):
        style = ExtendedStyle(
            base="base",
            name="child",
            font={"size": 14}
        ).extend(self.style)

        self.assertEqual(style.font.size, 14)
        self.assertEqual(style.font.bold, True)
        self.assertEqual(style.name, "child")

