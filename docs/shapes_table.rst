.. _shapes_table:

Tables
======

To create a table, use `createTableShape` method of slide.

Example:

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);

Rows
-------

A row is a child of a table. For creating a row, use `createRow` method of a Table shape.

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);
	$row = $tableShape->createRow();
	
Cells
-------
A cell is a child of a row.

You can access cell objects with `nextCell` method of a Row object.

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);
	$row = $tableShape->createRow();
	// Get the first cell
	$cellA1 = $row->nextCell();
	// Get the second cell
	$cellA2 = $row->nextCell();
	
You can access cell object directly.

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);
	$row = $tableShape->createRow();
	// Get the first cell
	$cellA1 = $row->getCell(0);
	// Get the second cell
	$cellA2 = $row->getCell(1);


Define margins of a cell
~~~~~~~~~~~~~~~~~~~~~~~~
Margins of cells are defined by margins of the first paragraph of cell.
Margins of cells are defined in pixels.

For defining margins of cell, you can use the `setMargin*` method of a Alignment object of the active paragraph of a Cell object.

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);
	$row = $tableShape->createRow();
	$cellA1 = $row->nextCell();
	$cellA1->getActiveParagraph()->getAlignment()
	    ->setMarginBottom(20)
	    ->setMarginLeft(40)
	    ->setMarginRight(60)
	    ->setMarginTop(80);


Define the text direction of a cell
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For defining the text direction of cell, you can use the `setTextDirection` method of the `getAlignment` method of a Cell object.
The width is in pixels.

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);
	$row = $tableShape->createRow();
	$cellA1 = $row->nextCell();
	$cellA1->getAlignment()->setTextDirection(\PhpOffice\PhpPresentation\Style\Alignment::TEXT_DIRECTION_VERTICAL_270);


Define the width of a cell
~~~~~~~~~~~~~~~~~~~~~~~~~~
The width of cells are defined by the width of cell of the first row.
If not defined, all cells widths are calculated from the width of the shape and the number of columns.

For defining the width of cell, you can use the `setWidth` method of a Cell object.
The width is in pixels.

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);
	$row = $tableShape->createRow();
	$cellA1 = $row->nextCell();
	$cellA1->setWidth(100);