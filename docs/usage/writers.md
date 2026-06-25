# Writers

## HTML
The name of the writer is `HTML`.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'HTML');
$writer->save(__DIR__ . '/sample.html');
```

## ODPresentation
The name of the writer is `ODPresentation`.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'PowerPoint2007');
$writer->save(__DIR__ . '/sample.pptx');
```

## PDF
The name of the writer is `PDF`.

``` php
<?php

use PhpOffice\PhpPresentation\Writer\PDF\DomPDF;

$writer = IOFactory::createWriter($oPhpPresentation, 'PDF');
$writer->setPDFAdapter(new DomPDF());
$writer->save(__DIR__ . '/sample.pdf');
```

## PowerPoint2007
The name of the writer is `PowerPoint2007`.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'PowerPoint2007');
$writer->save(__DIR__ . '/sample.pptx');
```

You can change the ZIP Adapter for the writer. By default, the ZIP Adapter is `ZipArchiveAdapter`.

``` php
<?php

use PhpOffice\Common\Adapter\Zip\PclZipAdapter;
use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;

$writer = IOFactory::createWriter($oPhpPresentation, 'PowerPoint2007');
$writer->setZipAdapter(new PclZipAdapter());
$writer->save(__DIR__ . '/sample.pptx');
```

## Keynote
The name of the writer is `Keynote`.

The writer produces a `.key` package following the legacy iWork '09 structure
(an `index.apxl` document). It writes slides, text placeholders (paragraphs and
runs), shape geometry and speaker notes. The output is round-trippable with the
`Keynote` reader; opening it directly in a current Keynote.app is not guaranteed
because of the format generation gap.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'Keynote');
$writer->save(__DIR__ . '/sample.key');
```

## Serialized
The name of the writer is `Serialized`.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'Serialized');
$writer->save(__DIR__ . '/sample.phppt');
```

You can change the ZIP Adapter for the writer. By default, the ZIP Adapter is `ZipArchiveAdapter`.

``` php
<?php

use PhpOffice\Common\Adapter\Zip\PclZipAdapter;
use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;

$writer = IOFactory::createWriter($oPhpPresentation, 'Serialized');
$writer->setZipAdapter(new PclZipAdapter());
$writer->save(__DIR__ . '/sample.phppt');
```
