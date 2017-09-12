 .. _templated_worksheet:

==================
TemplatedWorksheet
==================
The TemplatedWorksheet describes a a sheet in an excel file and is the bare bone used for building useful sheet templates such as the :ref:`TableSheet <TableSheet>`. A TemplatedWorksheet  is defined by following attributes:

    * It's ``sheetname`` which is used for identifying the sheets in the excel file
    * It's ``read()`` method which when implemented should return the relevant data contained in the sheet
    * It's ``write(data)`` method which when implemented should write the provided data to the sheet and make sure that it is properly formatted

The TemplatedWorksheet will handle managing the openpyxl worksheet so you do not have to worry about whether the sheet is created or not before you start writing.


To create a TemplatedWorksheet you should implement the read and write method. We'll demonstrate this by creating a DictSheet a TemplatedWorksheet which reads and writes simple key value pairs in a python dict.

.. literalinclude:: exampels/templated_worksheet.py
    :lines: 5-19

We can now add our DictSheet to a TemplatedWorkbook and use it to create an excel file.

.. literalinclude:: exampels/templated_worksheet.py
    :lines: 22-34

We can use the same TemplatedWorkbook to read the data from the file we just created.

.. literalinclude:: exampels/templated_worksheet.py
    :lines: 36-40

Have a look at the :ref:`TableSheet <TableSheet>` for a more advanced example. It includes type handling and plenty of styling.