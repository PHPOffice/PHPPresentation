# AutoShape

!!! warning
    Available only on the PowerPoint2007 Writer and Reader

The PowerPoint2007 Reader reads a shape as an `AutoShape` when it draws a preset other than a
rectangle and is neither a text box nor a placeholder. The text of an `AutoShape` is a plain string,
so a rectangle, where PowerPoint keeps most of its formatted text, is read as a `RichText`.

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