==============================
Welcome to openpyxl-templates!
==============================

Openpyxl-templates is an extension to `openpyxl <http://openpyxl.readthedocs.io/>`_ which is intended to simplify reading and writing of excel tables by limiting file layout. Openpyxl-templates templates does this by allowing users to specify which sheets to look for and which columns the tables on these sheets contains. This allows the user to provide data types and styling for each column which enables openpyxl-templates to validate when reading and formatting when writing data. The package also provides useful formatting shortcuts for freeze panes, hiding excess columns and formatting written data as a sortable table. 

The styling functionality is based on `Named Styles <http://openpyxl.readthedocs.io/en/default/styles.html#creating-a-named-style>`_ which provides a significant performance boost compared with cell styles as well as making them available for reuse in excel once exported.

----------
Fileformat
----------

Openpyxl-templates recognises the following elements in a excel sheet:
 * Sheets - Individual sheets identified by the sheetname
 * Tile - An optional titel found  in cell A1
 * Description - An optional description found directly below the title
 * Headers - The headers of the data table
 * Rows - Everything below the headers, when reading openpyxl-templates will read until the last populated row.

