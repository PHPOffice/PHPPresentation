# ![PHPPresentation](https://raw.githubusercontent.com/mvargasmoran/PHPPresentation/develop/docs/images/PHPPresentationLogo.png "PHPPresentation")

[![Latest Stable Version](https://poser.pugx.org/phpoffice/phppresentation/v)](https://packagist.org/packages/phpoffice/phppresentation)
[![Coverage Status](https://coveralls.io/repos/github/PHPOffice/PHPPresentation/badge.svg?branch=master)](https://coveralls.io/github/PHPOffice/PHPPresentation?branch=master)
[![Total Downloads](https://poser.pugx.org/phpoffice/phppresentation/downloads)](https://packagist.org/packages/phpoffice/phppresentation)
[![License](https://poser.pugx.org/phpoffice/phppresentation/license)](https://packagist.org/packages/phpoffice/phppresentation)

Branch Master : [![PHPPresentation](https://github.com/PHPOffice/PHPPresentation/actions/workflows/php.yml/badge.svg?branch=master)](https://github.com/PHPOffice/PHPPresentation/actions/workflows/php.yml)

PHPPresentation is a library written in pure PHP that provides a set of classes to write to different presentation file formats, i.e. Microsoft [Office Open XML](http://en.wikipedia.org/wiki/Office_Open_XML) (OOXML or OpenXML) or OASIS [Open Document Format for Office Applications](http://en.wikipedia.org/wiki/OpenDocument) (OpenDocument or ODF).

PHPPresentation is an open source project licensed under the terms of [LGPL version 3](https://github.com/PHPOffice/PHPPresentation/blob/develop/COPYING.LESSER). PHPPresentation is aimed to be a high quality software product by incorporating [continuous integration](https://github.com/PHPOffice/PHPPresentation/actions/workflows/php.yml) and [unit testing](https://coveralls.io/github/PHPOffice/PHPPresentation). You can learn more about PHPPresentation by reading the [Developers' Documentation](https://phpoffice.github.io/PHPPresentation) and the [API Documentation](https://phpoffice.github.io/PHPPresentation/docs/).

Read more about PHPPresentation:

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Getting started](#getting-started)
- [Contributing](#contributing)
- [Developers' Documentation](https://phpoffice.github.io/PHPPresentation/)
- [API Documentation](https://phpoffice.github.io/PHPPresentation/docs/)

### Features

- Create an in-memory presentation representation
- Set presentation meta data (author, title, description, etc)
- Add slides from scratch or from existing one
- Supports different fonts and font styles
- Supports different formatting, styles, fills, gradients
- Supports hyperlinks and rich-text strings
- Add images with different styles (positioning, rotation, shadow)
- Set printing options (header, footer, page margins, paper size, orientation)
- Set transitions between slides
- Output to different file formats: PowerPoint 2007 (.pptx), OpenDocument Presentation (.odp), Serialized Presentation)
- ... and lots of other things!

### Requirements

PHPPresentation requires the following:

- PHP 7.1+ 
- [ZIP Extension](http://php.net/manual/en/book.zip.php)
- [XML Parser Extension](http://www.php.net/manual/en/xml.installation.php)
- [XMLWriter Extension](http://php.net/manual/en/book.xmlwriter.php) (optional, used to write DOCX and ODT)
- [GD Extension](http://php.net/manual/en/book.image.php)

### Installation

#### Composer method

It is recommended that you install the PHPPresentation library [through composer](http://getcomposer.org/). To do so, add
the following lines to your ``composer.json``.

```json
{
    "require": {
       "phpoffice/phppresentation": "dev-master"
    }
}
```

#### Manual download method

Alternatively, you can download the latest release from the [releases page](https://github.com/PHPOffice/PHPPresentation/releases).
In this case, you will have to register the autoloader.
(Register autoloading is required only if you do not use composer in your project.)

```php
require_once 'path/to/PhpPresentation/src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();
```

You will also need to download the latest PHPOffice/Common release from its [releases page](https://github.com/PHPOffice/Common/releases).
And you will also have to register its autoloader, too.

```php
require_once 'path/to/PhpOffice/Common/src/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();
```

## Getting started

The following is a basic usage example of the PHPPresentation library.

```php
// with your own install
require_once 'src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();
require_once 'src/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();

// with Composer
require_once 'vendor/autoload.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;

$objPHPPowerPoint = new PhpPresentation();

// Create slide
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Create a shape (drawing)
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPresentation logo')
      ->setDescription('PHPPresentation logo')
      ->setPath('./resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shape->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);

// Create a shape (text)
$shape = $currentSlide->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
                   ->setSize(60)
                   ->setColor( new Color( 'FFE06B20' ) );

$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
$oWriterODP = IOFactory::createWriter($objPHPPowerPoint, 'ODPresentation');
$oWriterODP->save(__DIR__ . "/sample.odp");
```

More examples are provided in the [samples folder](samples/). You can also read the [Developers' Documentation](https://phpoffice.github.io/PHPPresentation/) and the [API Documentation](https://phpoffice.github.io/PHPPresentation/docs/) for more detail.

## Contributing

We welcome everyone to contribute to PHPPresentation. Below are some of the things that you can do to contribute:

- [Fork us](https://github.com/PHPOffice/PHPPresentation/fork) and [request a pull](https://github.com/PHPOffice/PHPPresentation/pulls) to the [master](https://github.com/PHPOffice/PHPPresentation) branch
- Submit [bug reports or feature requests](https://github.com/PHPOffice/PHPPresentation/issues) to GitHub
