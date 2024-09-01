# Introduction

Slides are pages in a presentation. Slides are stored as a zero based array in `PHPPresentation` object. 

## Create slide

Use the method `createSlide` to create a new slide and retrieve the slide for other operation such as creating shapes for that slide. The slide will be added at the end of slides collection.

``` php
<?php

$slide = $presentation->createSlide();
```

## Add slide to a specific position

Use the method `addSlide` to add an existing slide to a specific position. Without the parameter `$position`, it will be added at the end of slides collection.

``` php
<?php

use PhpOffice\PhpPresentation\Slide;

$slide = new Slide($presentation);
## Add it before all slides
$presentation->addSlide($slide, 0);
## Add it to position 1
$presentation->addSlide($slide, 1);
## Add it after all slides
$presentation->addSlide($slide);
```

## Properties

### Name

By default, a slide has not a name.
You can define it with the method `setName`.

``` php
<?php

$slide = $presentation->createSlide();
$slide->setName('Title of the slide');
```

### Visibility

By default, a slide is visible.
You can define it with the method `setIsVisible`.


``` php
<?php

$slide = $presentation->createSlide();
$slide->setIsVisible(false);
var_dump($slide->isVisible());
```