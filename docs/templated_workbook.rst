=================
TemplatedWorkbook
=================

The ``TemplatedWorkbook`` is our representation of the excel file and describes the file as a whole. To create a TemplatedWorkbook extend the base class and declare which sheets it includes using TemplatedWorksheets.


.. literalinclude:: examples/templated_workbook.py
    :lines: 1-6


To use your template to generate new excel files simply create an instance...


.. literalinclude:: examples/templated_workbook.py
    :lines: 9


... or provide it with a filename (or a file) to read an existing one.


.. literalinclude:: examples/templated_workbook.py
    :lines: 11


The TemplatedWorkbook will find all sheets which correspond to a TemplatedWorksheet. Once identified the TemplatedWorksheets can be used to interact with the underlying excel sheets. The matching is done based on the sheetname. The TemplatedWorkbook keeps track of the declaration order of the TemplatedWorksheets which enables it to make sure the the sheets are always in the correct order once the file has been saved. The identified sheets can also be iterated as illustrated below.


.. literalinclude:: examples/templated_workbook.py
    :lines: 14-15


To save the workbook simply call the save method and provide a filename


.. literalinclude:: examples/templated_workbook.py
    :lines: 18


