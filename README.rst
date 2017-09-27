.. image:: https://badge.fury.io/py/openpyxl-templates.svg
    :target: https://badge.fury.io/py/openpyxl-templates

==============================
Welcome to openpyxl-templates!
==============================

Openpyxl-templates is an extention to `openpyxl <http://openpyxl.readthedocs.io/>`_ which simplifies reading and writing excelfiles by formalizing their structure into templates. The package has two main components:

    1. The ``TemplatedWorkbook`` which describe the excel file
    2. The ``TemplatedSheets`` which describe the individual sheets within the file

The package is build for developers to be able to implement their own templates but also provides one useful templated sheet. The ``TableSheet`` which makes it significantly easier to read and write data from excel data tabels such as this one:

.. image:: examples/fruit_lovers.png

The TableSheet provides an easy way of defining columns and handles both styling and conversion to and from excel. See :ref:`quick start <quick_start>` for a demo.

Github
    https://github.com/SverkerSbrg/openpyxl-templates

Documentation
    http://openpyxl-templates.readthedocs.io/en/latest/

pypi
    https://pypi.python.org/pypi/openpyxl-templates

openpyxl
    https://openpyxl.readthedocs.io/en/default/


If you have any questions or ideas regarding the package feel free to reach out to me via GitHub.


.. warning::

    This package is still in beta. The api may still be subject to change and the documentation is patchy.