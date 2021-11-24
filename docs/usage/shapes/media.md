# Media

To create a video, create an object `Media`.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Media;

$media = new Media();
$media->setPath('file.wmv');
// $media->setPath('file.ogv');
$slide->addShape($media);
```

You can define text and date with setters.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Media;

$media = new Media();
$media->setName('Name of the Media');
$slide->addShape($media);
```

## Quirks

For Windows readers, the prefered file format is WMV.
For Linux readers, the prefered file format is OGV.