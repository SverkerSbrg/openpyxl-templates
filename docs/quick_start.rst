===========
Quick start
===========

Installation
------------

Install openpyxl-templates using pypi::

    pip install openpyxl-templates

Creating your template
----------------------
The first thing you need to do is to create your workbook template. Below is an example using the TableSheet to create an template for listing fruit loving persons.

.. literalinclude:: examples/simple_usage.py
    :lines: 1-26


Writing
-------
To write create an instance of your templated workbook, supply data to the sheets and save.

.. literalinclude:: examples/simple_usage.py
    :lines: 31-38

Openpyxl-templates handles all formatting. The code above produces the following excel sheet.

.. image:: examples/fruit_lovers.png

Reading
-------
To utilize the openpyxl-templates to read from an existing excel file, initialize your TemplatedWorkbook with a file (or a path to a file). Using the read method or simply itterating over a sheet will give you access to the data as namedtupels.

.. literalinclude:: examples/simple_usage.py
    :lines: 44-47

