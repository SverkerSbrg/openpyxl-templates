.. _quick_start:

===========
Quick start
===========

Installation
------------

Install openpyxl-templates using pypi::

    pip install openpyxl-templates

Creating your template
----------------------
The first we create our TemplatedWorkbook, which describes the structure of our file using TemplatedWorksheets. This template can then be used for both creating new files or reading existing ones. Below is an example using the TableSheet (a TemplatedWorksheet) to describe a excel file of people and their favorite fruits.

.. literalinclude:: examples/simple_usage.py
    :lines: 1-26


Writing
-------
To write create an instance of your templated workbook, supply data to the sheets and save to a file.

.. literalinclude:: examples/simple_usage.py
    :lines: 31-38

The TableSheet in this case will handle all formatting and produce the following sheet.

.. image:: examples/fruit_lovers.png

Reading
-------
To utilize the openpyxl-templates to read from an existing excel file, initialize your TemplatedWorkbook with a file (or a path to a file). Using the read method or simply itterating over a sheet will give you access to the data as namedtupels.

.. literalinclude:: examples/simple_usage.py
    :lines: 44-47

