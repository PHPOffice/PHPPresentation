# AutoShape

!!! warning
    Available only on the PowerPoint2007 Writer

To create a geometric form, create an object `AutoShape` and add it to slide.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\AutoShape;

$shape = new AutoShape();
$slide->addShape($shape)
```

Or we can use `createAutoShape` method of slide.

Example:

```php
$slide->createAutoShape();
```

## Type

An `AutoShape` is drawn as a heart until it is given a type. The types are the `AutoShape::TYPE_*`
constants, one per preset shape of PowerPoint.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\AutoShape;

$shape = new AutoShape();
$shape->setType(AutoShape::TYPE_ROUNDED_RECTANGLE)
    ->setOffsetX(100)
    ->setOffsetY(100)
    ->setWidth(300)
    ->setHeight(150);
```

## Fill and outline

The inside of the shape is its fill, as for every shape. Its line is its outline: a width in pixels
and a fill of its own. An `AutoShape` has no line until its outline is given a fill.

``` php
<?php

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

$shape->getFill()
    ->setFillType(Fill::FILL_SOLID)
    ->setStartColor(new Color('FFDDEEFF'));
$shape->getOutline()
    ->setWidth(3)
    ->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->setStartColor(new Color('FF1F4E79'));
```

!!! note
    `getBorder()`, which every shape has, is not how an `AutoShape` draws its line: set the
    outline instead.

## Rounded corner

The corner of a rounded rectangle is set in pixels with `setRoundRectCorner`. It is at most half
the shorter side of the shape; a shape given no corner is drawn with the corner of its preset.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\AutoShape;

$shape->setType(AutoShape::TYPE_ROUNDED_RECTANGLE)
    ->setRoundRectCorner(20);
```

The shadow, the hyperlink, the name and the alternative text of an `AutoShape` are set as for every
shape: see [Shapes](introduction.md).

## Text

You can define text of the geometric form with `setText` method.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\AutoShape;

$shape = new AutoShape();
// Define the text
$shape->setText('ABC');
// Return the text
$shape->getText();
```