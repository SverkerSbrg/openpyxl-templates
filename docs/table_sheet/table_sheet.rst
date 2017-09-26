.. _TableSheet:

==========
TableSheet
==========
The ``TableSheet`` is a ``TemplatedWorksheet`` making it easy for reading and write sheets with excel Data Tables. It is made up of an ordered set of typed columns which support when converting to and from Excel. Read more about what the columns do :ref:`here <table_sheet_columns>`.


Elements of the TableSheet
--------------------------

The ``TableSheet`` recognizes the following elements:

 * **Title** (optional) - A bold header for the Data Table
 * **Description** (optional) - A smaller description intended for simple instructions
 * **Columns** - The Columns in the datatable, which in turn is made up of **headers** and **rows**

.. literalinclude:: ../examples/table_sheet_elements.py

.. image:: ../examples/table_sheet_elements.png

The TableSheet does not support reading the title or description.

Creating the TableSheet
-----------------------

A ``TableSheet`` is created by extending the TableSheet class, declaring columns and optionally changing styling and other settings. Once the TableSheet class has been created an instance of this class is be supplied to the ``TemplatedWorkbook``.

Declaring columns
^^^^^^^^^^^^^^^^^
The columns are declared as class variables on a TableSheet which will identify and register the columns (in order of declaration). The columns are avaliable under the ``columns`` attribute.

A TableSheet must always have atleast one ``TableColumn``.

.. literalinclude:: ../examples/table_sheet.py
    :lines: 1-7

The column declaration supports inheritance, the following declaration is perfectly legal.

.. literalinclude:: ../examples/table_sheet.py
    :lines: 15-16

Note that the columns of the parent class are always considered to have been declared before the columns of the child.


All columns must have a header and there must not be any duplicated headers within the same sheet. The TableSheet will automatically use the attribute name used when declaring the column as header.

.. literalinclude:: ../examples/table_sheet.py
    :lines: 23-25

An instance of a column should never be used on multiple sheets.


TODO: Dynamic columns
^^^^^^^^^^^^^^^^^^^^^
TODO: Describe the ability to pass additional columns to __init__ via ``columns=[...]``


Writing
-------
Simple usage
^^^^^^^^^^^^

Writing is done by an iterable of objects to the write function and optionally a title and/or description. The write function will then:

    * Prepare the workbook by registering all required styles, data validation etc.
    * Write title, and description if they are supplied
    * Create the headers and rows
    * Apply sheet level formatting such as creating the Data Table and setting the freeze pane

Writing will always recreate the entire sheet from scratch, so any preexisting data will be lost. If you want to preserve your data you could read existing rows and combine them with the new data.

.. literalinclude:: ../examples/table_sheet_write_read.py
    :lines: 6-34

Using objects
^^^^^^^^^^^^^

The write accepts rows iterable containing tuples or list as in the example above. If an other type is encountered the columns will try to get the attribute directly from the object using ``getattr(object, column.object_attribute)``. The object_attribute can be defined explicitly and will default to the attribute name used when adding the column to the sheet.

.. literalinclude:: ../examples/table_sheet_write_read.py
    :lines: 39-52

Styling
^^^^^^^

The TableSheet has two style attributes:

    * ``tile_style`` - Name of the style to be used for the title, defaults to *"Title"*
    * ``description_style`` - Name of the style to be used for the description, defaults to *"Description"*

.. literalinclude:: ../examples/table_sheet_write_read.py
    :lines: 55-61

Styling of columns done on the columns themselves.

Make sure that the styles referenced are available either in the workbook or in the ``StyleSet`` of the ``TemplatedWorkbook``. Read more about styling :ref:`styling <here>`.

Additional settings
^^^^^^^^^^^^^^^^^^^
The write behaviour of the TableSheet can be modified with the following settings:
    * ``format_as_table`` - Controlling whether the TableSheet will format the output as a DataTable, defaults to *True*
    * ``freeze_pane`` - Controlling whether the TableSheet will utilize the freeze pane feature, defaults to *True*
    * ``hide_excess_columns`` - When enabled the TableSheet will hide all columns not used by columns, defaults to *True*


Reading
-------

Simple usage
^^^^^^^^^^^^

The ``read`` method does two things. First it will verify the format of the file by looking for the header row. If the headers cannot be found a en exception will be raised. Once the headers has been found all subsequent rows in the excel will be treated as data and parsed to `namedtuples <https://docs.python.org/3/library/collections.html#collections.namedtuple>`_ automatically after the columns has transformed the data from excel to python.

.. literalinclude:: ../examples/table_sheet_write_read.py
    :lines: 65-67

Iterate directly
^^^^^^^^^^^^^^^^

The TableSheet can also be used as an iterator directly

.. literalinclude:: ../examples/table_sheet_write_read.py
    :lines: 71-72

Exception handling
^^^^^^^^^^^^^^^^^^
The way the TableSheet handles exceptions can be configured by setting the ``exception_policy``. It can be set on the TableSheet class or passed as an argument to the read function. The following policies are avaliable:
    * ``RaiseCellException`` (default) - All exceptions will be raised when encountered
    * ``RaiseRowException`` - Cell level exceptions such as type errors, in the same row will be collected and raised as a RowException
    * ``RaiseSheetException`` - All row exceptions will be collected and raised once reading has finished. So that all valid rows will be read, and all exceptions will be recorded.
    * ``IgnoreRow`` - Invalid rows will be ignored

The policy only applies to exceptions occuring when reading rows. Exceptions such as ``HeadersNotFound`` will be raised irregardless.


Reading without looking for headers
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Looking for headers can be disabled by setting ``look_for_headers`` to *False* or passing it as a named argument directly to the read function. When this is done the TableSheet will start looking for valid rows at once. This will most likely cause an exception if the title, description or header row is present since they will be treated as rows.

Customization
-------------
The TableSheet is built with customization in mind. If you want your table to yield something else then a ``namedtuple`` for each row. It is easy to achieve by overriding the ``create_object`` method.

.. literalinclude:: ../examples/customization.py
    :lines: 6-23

Feel free to explore the source code for additional possibilities. If you are missing hook or add a feature useful for others, feel free to submit a push request.