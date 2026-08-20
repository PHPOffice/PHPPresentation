# Tables
To create a table, use `createTableShape` method of slide.

Example:

``` php
<?php

$tableShape = $slide->createTableShape($columns);
```

## Header row and banded rows

A table is written with its first row styled as a header row and with alternating row bands.
Both can be turned off with `setFirstRow` and `setBandRow`.

A header row is announced as a header by assistive technologies, so a table whose first row holds
data rather than column labels — a single-column list, for instance — should turn it off, or its
first entry is read out as the heading of the column.

``` php
<?php

$tableShape = $slide->createTableShape($columns);
$tableShape->setFirstRow(false);
$tableShape->setBandRow(false);
```

Each writer says it in the terms of its own format: the PowerPoint2007 Writer sets `firstRow` on the
table properties, the ODPresentation Writer wraps the first row in `table:table-header-rows`.

!!! warning
    `setBandRow` is available only on the PowerPoint2007 Writer. ODF has no equivalent — banding is
    a matter of styling each row there.

## Rows
A row is a child of a table. For creating a row, use `createRow` method of a Table shape.

``` php
<?php

$tableShape = $slide->createTableShape($columns);
$row = $tableShape->createRow();
```

### Define the fill

A fill set on a row paints every cell of it, which is how a table gets a banded look without
setting the same fill on each cell in turn. A cell that asks for a fill of its own keeps it: the
fill of the row is what a cell falls back to, not what it is overruled by.

``` php
<?php

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

$row = $tableShape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
```

Neither format carries a fill on a table row -- in OOXML a row has nothing but a height, and the
row style of ODF takes a flat colour that LibreOffice does not draw -- so both Writers write the
fill of the row on each of its cells.

## Cells
A cell is a child of a row.

You can access cell objects with `nextCell` method of a Row object.

``` php
<?php

$tableShape = $slide->createTableShape($columns);
$row = $tableShape->createRow();
// Get the first cell
$cellA1 = $row->nextCell();
// Get the second cell
$cellA2 = $row->nextCell();
```

You can access cell object directly.

``` php
<?php

$tableShape = $slide->createTableShape($columns);
$row = $tableShape->createRow();
// Get the first cell
$cellA1 = $row->getCell(0);
// Get the second cell
$cellA2 = $row->getCell(1);
```

### Define margins
Margins of cells are defined by margins of the first paragraph of cell.
Margins of cells are defined in pixels.

For defining margins of cell, you can use the `setMargin*` method of a Alignment object of the active paragraph of a Cell object.

``` php
<?php

$tableShape = $slide->createTableShape($columns);
$row = $tableShape->createRow();
$cellA1 = $row->nextCell();
$cellA1->getActiveParagraph()->getAlignment()
    ->setMarginBottom(20)
    ->setMarginLeft(40)
    ->setMarginRight(60)
    ->setMarginTop(80);
```

### Define the text direction
For defining the text direction of cell, you can use the `setTextDirection` method of the `getAlignment` method of a Cell object.
The width is in pixels.

``` php
<?php

use PhpOffice\PhpPresentation\Style\Alignment;

$tableShape = $slide->createTableShape($columns);
$row = $tableShape->createRow();
$cellA1 = $row->nextCell();
$cellA1->getAlignment()->setTextDirection(Alignment::TEXT_DIRECTION_VERTICAL_270);
```

### Define the width
The width of cells are defined by the width of cell of the first row.
If not defined, all cells widths are calculated from the width of the shape and the number of columns.

For defining the width of cell, you can use the `setWidth` method of a Cell object.
The width is in pixels.

``` php
<?php

$tableShape = $slide->createTableShape($columns);
$row = $tableShape->createRow();
$cellA1 = $row->nextCell();
$cellA1->setWidth(100);
```