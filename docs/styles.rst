.. _styling:

===================
Working with styles
===================

Styling in openpyxl-templates is entirely reliant on the ``NamedStyle`` provided by openpyxl. Using NamedStyles offers significantly better performance compared to styling each individual cells, as well as the benefits of making the styles avaliable in the produced excel file.

To manage all named styles within a TemplatedWorkbook openpyxl-templates uses the ``StyleSet`` which is a dictionary like collection of NamedStyles. Using the StyleSet when creating templates is entierly optional but offers several advantages:

 - Using a common colleciton of styles for all TemplatedSheets makes it easier to avoid duplicated styles and name conflicts
 - The StyleSet accepts ``ExtendedStyle`` s as well as NamedStyles which enables inheritance between styles within the StyleSet
 - Styles will only be added to the excel file when they are needed allowing the developer to use a single StyleSet for multiple templates without having to worry about unused styles being included in the excel file.
 - Using NamedStyles offers significantly better performace compared to styling each cell individually when writing a large amout of data

-------------------
Creating a StyleSet
-------------------

To create a StyleSet simply pass NamedStyles or ExtendedStyles as arguments.

.. literalinclude:: examples/styles.py
    :lines: 6-22

If we want to avoid having to redeclare the font we could refactor the above example using an ExtendedStyle

.. literalinclude:: examples/styles.py
    :lines: 24-39

The ExtendedStyle can be viewed as a NamedStyle factory. It accepts the same arguments as a NamedStyle with the addition of the ``base`` which is the name of the intended parent. The StyleSet will make sure that the parent style is found irregardless of declaration order.

The font, border, alignment and fill arguments of the NamedStyle are supplied as objects which (since they have default values) prevent inheritance. To circomvent this limitation you can supply the kwarg dicts to the extended style instead of the object itself, as we have done in the example above. If the openpyxl class is used instead of kwargs, inheritance will be broken

.. literalinclude:: examples/styles.py
    :lines: 41-56

if a name is declared multiple times the last declaration will take precedence making it easy to modify an existing StyleSet. One common usage for this is to modify the DefaultStyleSet. Which is demonstrated :ref:`ModifyDefaultStyleSet <below>`.


----------------
Accessing styles
----------------

TODO

.. _DefaultStyleSet:

---------------
DefaultStyleSet
---------------

Openpyxl-templates includes a ``DefaultStyleSet`` which is used as a fallback for all TemplatedWorkbook. Many of the styles it declares (or their names) are required by the ``TableSheet``. The DefaultStyleSet is defined like this

.. literalinclude:: ../openpyxl_templates/styles.py
    :lines: 114-184

.. _ModifyDefaultStyleSet:

If you which to modify the DefaultStyle you can easily replace any or all of the styles it contains by passing them as arguments to the constructor. Below we change the fill color of all headers by replacing the "Header" ExtendedStyle.

 TODO: Example
