# Readers

## ODPresentation
The name of the reader is `ODPresentation`.

``` php
<?php

$reader = IOFactory::createReader('ODPresentation');
$reader->load(__DIR__ . '/sample.odp');
```

## PowerPoint97
The name of the reader is `PowerPoint97`.

``` php
<?php

$reader = IOFactory::createReader('PowerPoint97');
$reader->load(__DIR__ . '/sample.ppt');
```

## PowerPoint2007
The name of the reader is `PowerPoint2007`.

``` php
<?php

$reader = IOFactory::createReader('PowerPoint2007');
$reader->load(__DIR__ . '/sample.pptx');
```

## Serialized
The name of the reader is `Serialized`.

``` php
<?php

$reader = IOFactory::createReader('Serialized');
$reader->load(__DIR__ . '/sample.phppt');
```
