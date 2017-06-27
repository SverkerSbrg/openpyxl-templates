==========
TableSheet
==========
The ``TableSheet`` is a ``TemplatedWorksheet`` making it easy for reading and write excel Data Table. It is made up of an ordered set of typed columns which support when converting to and from Excel. Read more about what the columns do :ref:`here <table_sheet_columns>`.


Elements of the TableSheet
--------------------------

The ``TableSheet`` recognizes the following elements:
 * **Title** (optional) - A bold header for the Data Table
 * **Description** (optional) - A smaller description intended for simple instructions
 * **Columns** - The Columns in the datatable, which in turn is made up of **headers** and **rows**

.. literalinclude:: ../examples/table_sheet_elements.py

.. image:: ../examples/table_sheet_elements.png

The TableSheet does not support reading the title and description elements.

Configuration
-------------

Creating a ``TableSheet`` follows the same syntax as the ``TemplatedWorkbook``. The columns are declared as class variables on a TableSheet which will identify and register the columns. The complete and ordered, set of columns are accessible under the ``.columns`` attribute.

.. literalinclude:: ../examples/table_sheet.py
    :lines: 1-10

The column declaration has full support for inheritance. The following declaration is perfectly legal.

.. literalinclude:: ../examples/table_sheet.py
    :lines: 13-14

Note that the columns of the parent class are always considered to have been declared before the columns of the child.

A TableSheet must always have atleast one Column.

All columns must have a header and there must not be any duplicated headers within the same sheet. The TableSheet will automatically use the attribute name used when declaring the column as header.

.. literalinclude:: ../examples/table_sheet.py
    :lines: 17-19