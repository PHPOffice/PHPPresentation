# AutoShape

!!! warning
    Available on the PowerPoint2007 Writer, and on the ODPresentation Writer and Reader

The ODPresentation Writer writes an AutoShape the way LibreOffice writes an OOXML preset shape: a
`draw:custom-shape` of type `ooxml-<type>`, carrying the geometry it is drawn from, as OpenDocument
names no preset shapes. The ODPresentation Reader reads such a shape back, and a custom shape
LibreOffice draws with a geometry of its own as the preset its export to PowerPoint names.

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