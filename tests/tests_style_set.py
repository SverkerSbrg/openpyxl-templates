from unittest import TestCase

from openpyxl.styles import NamedStyle, Font

from openpyxl_templates.old.workbook import StyleSet, ExtendedStyle


class StyleSetTest(TestCase):
    def test_return_none(self):
        styles = StyleSet()

        for key in (
                "__default__",
                "__row__",
                "__title__",
                "__header__",
                "anything else",
        ):
            self.assertIsNone(styles[key])

    def test_style(self):
        style = NamedStyle(name="test", font=Font(bold=True))
        styles = StyleSet(style)
        self.assertIs(style, styles["test"])

    def test_default(self):
        default = NamedStyle(font=Font(bold=True))
        styles = StyleSet(__default__=default)

        for key in (
                "__default__",
                "__row__",
                "__title__",
                "__header__",
        ):
            self.assertIs(default, styles[key])

    def test_no_duplicated_keys(self):
        style1 = NamedStyle(name="test", font=Font(bold=True))
        style2 = NamedStyle(name="test", font=Font(bold=False))
        with self.assertRaises(Exception):
            styles = StyleSet(style1, style2)

    def test_no_duplicated_styles(self):
        style1 = NamedStyle(name="test", font=Font(bold=True))
        style2 = NamedStyle(name="test2", font=Font(bold=True))

        styles = StyleSet(style1, style2)
        self.assertIs(styles["test"], styles["test2"])

    def test_get_by_style(self):
        _styles = [
            NamedStyle(name="test", font=Font(bold=False)),
            NamedStyle(name="test2", font=Font(bold=True))
        ]

        styles = StyleSet(*_styles)

        for style in _styles:
            self.assertIs(styles[style], style)

    def test_extended_style(self):
        base = NamedStyle(name="base", font=Font(bold=True, size=14, italic=True))
        extended = ExtendedStyle(base="base", name="extended", font={"size": 20, "italic": False})

        styles = StyleSet(base, extended)

        self.assertEqual(
            styles["extended"],
            NamedStyle(name="extended", font=Font(bold=True, size=20, italic=False))
        )
