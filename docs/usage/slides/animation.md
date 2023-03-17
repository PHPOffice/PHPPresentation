# Animations


You can create multiples animations in a slide.


``` php
<?php
use PhpOffice\PhpPresentation\Slide\Animation;

$animation1 = new Animation();
$animation1->addShape($drawing);
$slide->addAnimation($animation1);

$animation2 = new Animation();
$animation2->addShape($richtext);
$slide->addAnimation($animation2);

```