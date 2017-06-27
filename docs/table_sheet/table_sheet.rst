==========
TableSheet
==========
The ``TableSheet`` is a ``TemplatedWorksheet`` making it easy for reading and write excel Data Table. It is made up of an ordered set of typed columns which support when converting to and from Excel. Read more about what the columns to <<HERE>>

The ``TableSheet`` recognizes the following elements:
 * Title (optional) - A bold header for the Data Table
 * Description (optional) - A smaller description intended for simple instructions
 * Columns - The Columns in the datatable, which in turn is made up of a row of headers and a number of data rows

.. literalinclude:: examples/table_sheet_elements.py

.. image:: ../examples/table_sheet_elements.png