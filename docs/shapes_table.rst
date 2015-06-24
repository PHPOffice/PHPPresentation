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
	
You can access cell object directly

.. code-block:: php

	$tableShape = $slide->createTableShape($columns);
	$row = $tableShape->createRow();
	// Get the first cell
	$cellA1 = $row->getCell(0);
	// Get the second cell
	$cellA2 = $row->getCell(1);
	

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