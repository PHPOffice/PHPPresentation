# Introduction

Slides are pages in a presentation. Slides are stored as a zero based array in `PHPPresentation` object. 

## Create slide

Use the method `createSlide` to create a new slide and retrieve the slide for other operation such as creating shapes for that slide. The slide will be added at the end of slides collection.

``` php
<?php

$slide = $presentation->createSlide();
```

## Add slide to a specific position

Use the method `addSlide` to add an existing slide to a specific position. Without the parameter `$index`, it will be added at the end of slides collection.

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

## Move a slide

Use the method `moveSlide` to move a slide that is already in the presentation to another position.

``` php
<?php

$presentation->moveSlide($slide, 2);
```

The slide is taken out before it is put back, so the index counts in the collection *without* it.
In a presentation of `A B C D`, moving `A` to 2 gives `B C A D`.

It throws an `OutOfBoundsException` for an index past the last slide, and an
`InvalidParameterException` for a slide the presentation does not hold.

## Properties

### Name

By default, a slide has not a name.
You can define it with the method `setName`.

``` php
<?php

$slide = $presentation->createSlide();
$slide->setName('Title of the slide');
```

It is written as the `name` attribute of `p:cSld` in PowerPoint2007 files and as the
`draw:name` attribute of `draw:page` in ODPresentation files. Both readers restore it.

### Background

By default, a slide has no background, and neither has the master slide it is based on: what shows
through is the light colour of the theme. You can define one with the method `setBackground`, on a
slide for that slide alone or on the master for every slide based on it.

``` php
<?php

use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Style\Color;

$background = new BackgroundColor();
$background->setColor(new Color('FFCC00'));

$presentation->getAllMasterSlides()[0]->setBackground($background);
```

A background is a fill of its own, drawn over the whole page. A converter that turns the file into
a tagged PDF has nothing to say about that fill, so it lands outside the tag tree -- which is why
one is written only when it is asked for.

Both Writers keep the background of a master: the PowerPoint2007 Writer as the `p:bg` of the master
slide, the ODPresentation Writer as the drawing-page style of the master page. A colour and an
image are both carried.

### Visibility

By default, a slide is visible.
You can define it with the method `setIsVisible`.


``` php
<?php

$slide = $presentation->createSlide();
$slide->setIsVisible(false);
var_dump($slide->isVisible());
```