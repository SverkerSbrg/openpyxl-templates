from unittest import TestCase

from openpyxl_templates.utils import OrderedType, class_property


class MagicString(str):
    pass


class OrderedTypeTestClass(object):
    __metaclass__ = OrderedType
    item_class = MagicString

    @class_property
    def items(self):
        return list(self._items.values())


class OrderedTypeTests(TestCase):
    def test_objects_identified(self):
        class Test(OrderedTypeTestClass):
            string1 = MagicString("string1")
            string2 = MagicString("string2")
            string3 = MagicString("string3")

        result = list(Test.items)
        for index, string in enumerate((Test.string1, Test.string2, Test.string3)):
            self.assertEqual(result[index], string)

    def test_inheritence(self):
        class Test(OrderedTypeTestClass):
            string1 = MagicString("string1")
            string2 = MagicString("string2")
            string3 = MagicString("string3")

        class Test2(Test):
            string2 = MagicString("new_string2")
            string4 = MagicString("string4")

        result = list(Test2.items)
        for index, string in enumerate((Test2.string1, Test2.string2, Test2.string3, Test2.string4)):
            self.assertEqual(result[index], string)

    def test_multiple_inheritence(self):
        class Parent1(OrderedTypeTestClass):
            string1 = MagicString("Parent1.string1")
            string2 = MagicString("Parent1.string2")

        class Parent2(OrderedTypeTestClass):
            string2 = MagicString("Parent2.string2")
            string3 = MagicString("Parend2.string3")

        class Child1(Parent2, Parent1):
            string3 = MagicString("child.string3")
            string4 = MagicString("child.string4")

        class Child2(Parent1, Parent2):
            string3 = MagicString("child.string3")
            string4 = MagicString("child.string4")

        result = Child1.items
        for index, attr in enumerate(["string1", "string2", "string3", "string4"]):
            self.assertEqual(result[index], getattr(Child1, attr))

        result = Child2.items
        for index, attr in enumerate(["string2", "string3", "string1", "string4"]):
            self.assertEqual(result[index], getattr(Child2, attr))

