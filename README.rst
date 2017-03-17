==============================
Welcome to openpyxl-templates!
==============================

Openpyxl-templates is an extension to `openpyxl <http://openpyxl.readthedocs.io/>`_ which is intended to simplify reading and writing of excel tables by limiting restricting the layout of the excel to a standardized table. Openpyxl-templates works based on a template for the file which specifying its strucutre and content. This template has tree levels the workbook, the worksheet and the data columns on each individual sheet. The columns allows for data validation and can ensure that the correct number format is used.

Openpyxl-templates also provides shortcuts to features common when working with these kind of files such as "format as table" and the ability to hide all columns right of the last column in a sheet.

Features
--------
* **Type support** provided by the column definitions enables more robust conversion to python when reading and removes the hassle of setting the number format yourself. For example openpyxl-templates can correctly read a date even when the number format is set to number rather then a date
* **Data validation**
* **Uniform styling** the workbook template allows you to set a style for the entire workbook, which reduces boiler plate.
   * Override styles for entire worksheets or single columns
   * Based on `Named Styles <http://openpyxl.readthedocs.io/en/default/styles.html#creating-a-named-style>`_ making them avaliable in the resulting excel as well as providing a performance boost compared with using cell styles
   * **Format-as-table** automatically
   * **Hide excess columns**


Fileformat
----------

Openpyxl-templates recognises the following elements in a excel sheet:

* Sheets - Individual sheets identified by the sheetname
* Title - An optional title found  in cell A1
* Description - An optional description found directly below the title
* Headers - The headers of the data table
* Rows - The actual data, everything below the headers

