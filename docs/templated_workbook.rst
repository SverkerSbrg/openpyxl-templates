=================
TemplatedWorkbook
=================

The ``TemplatedWorkbook`` is our representation of the excel file and consists of a collection of templated sheets.

Create a workbook template for your excel by extending the base class and declare templated worksheets as class variables.


.. literalinclude:: examples/templated_workbook.py
    :lines: 1-6


To create a new file simply create and instance


.. literalinclude:: examples/templated_workbook.py
    :lines: 9


Or use the filename argument to open an existing excel file.


.. literalinclude:: examples/templated_workbook.py
    :lines: 11


The templated workbook will identify all templated sheets and make them avaliable in order of declaration.


.. literalinclude:: examples/templated_workbook.py
    :lines: 14-15


To save the workbook to an .xlsx simply call the save method and provide a filename


.. literalinclude:: examples/templated_workbook.py
    :lines: 18


The templated workbook automatically


The worksheets will both be avaliable as a list (sorted in order of declaration) ``templated_workbook.templated_worksheets``.

