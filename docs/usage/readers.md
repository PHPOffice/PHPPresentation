# Readers

## Keynote
The name of the reader is `Keynote`.

``` php
<?php

$reader = IOFactory::createReader('Keynote');
$reader->load(__DIR__ . '/sample.key');
```

The reader reads the text of every slide, its speaker note and the images the slide uses, out of a
Keynote '13 (and later) package -- its `Index/*.iwa` components -- as well as out of the
`index.apxl` document of a Keynote '09 package. Anything else the format carries (styles, tables,
charts, transitions) is not read yet.

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\Keynote;

$reader = new Keynote();
$reader->load(__DIR__ . '/sample.key', Keynote::SKIP_IMAGES);
```

## ODPresentation
The name of the reader is `ODPresentation`.

``` php
<?php

$reader = IOFactory::createReader('ODPresentation');
$reader->load(__DIR__ . '/sample.odp');
```

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\ODPresentation;

$reader = new ODPresentation();
$reader->load(__DIR__ . '/sample.odp', ODPresentation::SKIP_IMAGES);
```

## PowerPoint97
The name of the reader is `PowerPoint97`.

``` php
<?php

$reader = IOFactory::createReader('PowerPoint97');
$reader->load(__DIR__ . '/sample.ppt');
```

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\PowerPoint97;

$reader = new PowerPoint97();
$reader->load(__DIR__ . '/sample.ppt', PowerPoint97::SKIP_IMAGES);
```

## PowerPoint2007
The name of the reader is `PowerPoint2007`.

``` php
<?php

$reader = IOFactory::createReader('PowerPoint2007');
$reader->load(__DIR__ . '/sample.pptx');
```

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\PowerPoint2007;

$reader = new PowerPoint2007();
$reader->load(__DIR__ . '/sample.pptx', PowerPoint2007::SKIP_IMAGES);
```

## Serialized
The name of the reader is `Serialized`.

``` php
<?php

$reader = IOFactory::createReader('Serialized');
$reader->load(__DIR__ . '/sample.phppt');
```
