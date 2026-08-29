# Writers

## HTML
The name of the writer is `HTML`.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'HTML');
$writer->save(__DIR__ . '/sample.html');
```

## Keynote
The name of the writer is `Keynote`.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'Keynote');
$writer->save(__DIR__ . '/sample.key');
```

The writer writes the text of every slide, its speaker note and the images it uses, as the
`index.apxl` package Keynote '09 reads. The `Index/*.iwa` shape of Keynote '13 and later is not
written.

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
