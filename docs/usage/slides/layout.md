# Slides Layout

A slide is a hierarchy of three components:

- The master slide upon which the slide is based : it specifies such properties as the font styles for the title, body, and footer, placeholder positions for text and objects, bullets styles, and background;
- The slide layout which is applied to the the slide : it permits to override what is specified in the master slide;
- The slide itself: it contains content and formatting that is not already specified by the master slide and the slide layout.

Placeholders permit to link these three components together in order that the override is possible.

## Master slides

You can access to all master slides with the method `getAllMasterSlides` or create one with `createMasterSlide`.

``` php
<?php

$arraySlideMasters = $presentation->getAllMasterSlides();
$masterSlide = $presentation->createMasterSlide();
```

## Slides Layout

You can access to all slide layout from a master with the method `getAllSlideLayouts` or create one with `createSlideLayout`.


``` php
<?php

$arraySlideLayouts = $masterSlide->getAllSlideLayouts();
$slideLayout = $masterSlide->createSlideLayout();
```

## Placeholders

For each master slide or slide layout, you can add any shape like on a slide.

``` php
<?php

$shape = $masterSlide->createChartShape();
$shape = $slideLayout->createTableShape();
```

You can define a shape as a placeholder for each level with the method `setPlaceHolder`.
A shape defined in each level will have an override for its formatting in each level.


``` php
<?php

use PhpOffice\PhpPresentation\Shape\Placeholder;
$shape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_TITLE));
```
