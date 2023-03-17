# Writers

## ODPresentation
The name of the writer is `ODPresentation`.

``` php
<?php

$writer = IOFactory::createWriter($oPhpPresentation, 'PowerPoint2007');
$writer->save(__DIR__ . '/sample.pptx');
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
