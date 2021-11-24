# Drawing

To create a drawing, you have multiples sources :

- Base64
- File
- GD
- ZipFile

You can add multiples formats of image :

- GIF
- JPEG
- PNG
- SVG

## File

To create a drawing, use `createDrawingShape` method of slide.

``` php
<?php

$shape = $slide->createDrawingShape();
$shape->setName('Unique name')
    ->setDescription('Description of the drawing')
    ->setPath('/path/to/drawing.filename');
```

It's an alias for :

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Drawing\File;

$shape = new File();
$shape->setName('Unique name')
    ->setDescription('Description of the drawing')
    ->setPath('/path/to/drawing.filename');
$slide->addShape($shape);
```

## Base64

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Drawing\Base64;

$shape = new Base64();
$shape->setName('Sample image')
    ->setDescription('Sample image')
    ->setImageResource($gdImage)
    ->setData('data:image/jpeg;base64,..........');
$slide->addShape($shape);
```

## GD

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Drawing\Gd;

$gdImage = imagecreatetruecolor($width, $height);

$shape = new Gd();
$shape->setName('Sample image')
    ->setDescription('Sample image')
    ->setImageResource($gdImage)
    ->setRenderingFunction(Drawing\Gd::RENDERING_JPEG)
    ->setMimeType(Drawing\Gd::MIMETYPE_DEFAULT);
$slide->addShape($shape);
```

## ZipFile

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Drawing\ZipFile;

$shape = new ZipFile();
$shape->setName('Sample image')
    ->setDescription('Sample image')
    ->setPath('zip://myzipfile.zip#path/in/zip/img.ext')
$slide->addShape($shape);
```
