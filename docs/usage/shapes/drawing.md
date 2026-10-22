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

A drawing that was given no path answers with an empty path: `getPath()` is typed `string`, and a drawing that was made rather than read has no path until one is set on it.

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

A GD drawing that was read out of a package keeps the bytes the file holds, and `getContents()` gives them back rather than an encoding of the picture made again. `setContents()` is how they are put there. Setting the resource drops them: they describe the resource they were decoded into, and another resource is not that one.

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
