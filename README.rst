==============================
Welcome to openpyxl-templates!
==============================

Openpyxl-templates is an extension to `openpyxl <http://openpyxl.readthedocs.io/>`_ which is intended to simplify reading and writing of excelfiles by formalizing their structure. The package is built around the ``TemplatedWorkbook``, a class-based representation of an excel file. The workook is described by a number of ``TemplatedWorksheet``s which describe the individual sheets. The TemplatedWorksheet is an abstract base class for classbased representations of individual sheets. It provides an interface for reading and writing (including styleing) sheets.

The ``TableSheet`` is (currently) the only implementation of the TemplatedSheet. It allows the developer to easily read and write Data Tables.

Github
    https://github.com/SverkerSbrg/openpyxl-templates

Documentation
    http://openpyxl-templates.readthedocs.io/en/latest/

pypi
    https://pypi.python.org/pypi/openpyxl-templates

openpyxl
    https://openpyxl.readthedocs.io/en/default/


.. warning::

    This package is still in beta. The api may still be subject to change and the documentation is patchy.