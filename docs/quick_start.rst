===========
Quick start
===========

Installation
------------

Install openpyxl-templates using pypi.

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

The templated workbook and its templated sheets handles all formatting. Here is the persons sheet created by the code above.

.. image:: examples/fruit_lovers.png

Reading
-------
To read use the filename argument of your workbook template.

.. literalinclude:: examples/simple_usage.py
    :lines: 44-47

