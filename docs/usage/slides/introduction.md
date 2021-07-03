# Introduction

Slides are pages in a presentation. Slides are stored as a zero based array in `PHPPresentation` object. Use the method `createSlide` to create a new slide and retrieve the slide for other operation such as creating shapes for that slide.

## Name

By default, a slide has not a name.
You can define it with the method `setName`.

``` php
<?php

$slide = $presentation->createSlide();
$slide->setName('Title of the slide');
```

## Visibility

By default, a slide is visible.
You can define it with the method `setIsVisible`.


``` php
<?php

$slide = $presentation->createSlide();
$slide->setIsVisible(false);
var_dump($slide->isVisible());
```