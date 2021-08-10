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