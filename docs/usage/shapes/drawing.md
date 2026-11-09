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
- WMF, EMF & EMF+ (Windows metafiles), if the optional library [phpoffice/wmf](https://github.com/PHPOffice/WMF) and the GD extension are installed

## Geometry

A drawing is drawn in a rectangle, and it can be drawn in another shape instead: the name of a
preset shape, which is what a picture cropped to a shape is written with.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\AutoShape;

$shape = $slide->createDrawingShape();
$shape->setPath('/path/to/drawing.filename');
$shape->setGeometry(AutoShape::TYPE_OVAL);
```

The presets are the ones of the `AutoShape` class, so `AutoShape::TYPE_OVAL` is `ellipse` and
`AutoShape::TYPE_HEART` is `heart`. The geometry a drawing is drawn in is the name of the shape
and nothing else: the picture keeps its own size and its own place, and the shape is drawn
around it.

The PowerPoint2007 Writer writes the geometry and its Reader reads it back. The ODPresentation
Writer has no geometry for a drawing and writes the picture as it is.

## Windows metafiles

To add a WMF, an EMF or an EMF+ file, install the optional library [phpoffice/wmf](https://github.com/PHPOffice/WMF) :

``` sh
composer require phpoffice/wmf
```

It says what a metafile is (`image/x-wmf` or `image/x-emf` ; an EMF+ file is an EMF file) and measures it, which `getimagesize()` can't do.

``` php
<?php

$shape = $slide->createDrawingShape();
$shape->setName('Chart')
    ->setDescription('A chart drawn as a metafile')
    ->setPath('/path/to/chart.emf');
```

The PowerPoint2007 and ODPresentation Writers store a metafile as it is, and their Readers read it back.
The Writers whose target can't show a metafile (HTML, PDF and Keynote) write the PNG it draws instead.

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
