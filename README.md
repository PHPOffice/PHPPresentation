# ![PHPPresentation](https://github.com/PHPOffice/PHPPresentation/raw/master/docs/images/PHPPowerPointLogo.png "PHPPresentation")

[![Latest Stable Version](https://poser.pugx.org/phpoffice/phppowerpoint/v/stable.png)](https://packagist.org/packages/phpoffice/phppowerpoint)
[![Build Status](https://travis-ci.org/PHPOffice/PHPPresentation.svg?branch=master)](https://travis-ci.org/PHPOffice/PHPPresentation)
[![Code Quality](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/badges/quality-score.png?s=b5997ce59ac2816b4514f3a38de9900f6d492c1d)](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/)
[![Code Coverage](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/badges/coverage.png?s=742a98745725c562955440edc8d2c39d7ff5ae25)](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/)
[![Total Downloads](https://poser.pugx.org/phpoffice/phppowerpoint/downloads.png)](https://packagist.org/packages/phpoffice/phppowerpoint)
[![License](https://poser.pugx.org/phpoffice/phppowerpoint/license.png)](https://packagist.org/packages/phpoffice/phppowerpoint)
[![Join the chat at https://gitter.im/PHPOffice/PHPPowerPoint](https://img.shields.io/badge/GITTER-join%20chat-green.svg)](https://gitter.im/PHPOffice/PHPPowerPoint)


PHPPresentation is a library written in pure PHP that provides a set of classes to write to different presentation file formats, i.e. Microsoft [Office Open XML](http://en.wikipedia.org/wiki/Office_Open_XML) (OOXML or OpenXML) or OASIS [Open Document Format for Office Applications](http://en.wikipedia.org/wiki/OpenDocument) (OpenDocument or ODF). 

PHPPresentation is an open source project licensed under the terms of [LGPL version 3](https://github.com/PHPOffice/PHPPresentation/blob/develop/COPYING.LESSER). PHPPresentation is aimed to be a high quality software product by incorporating [continuous integration](https://travis-ci.org/PHPOffice/PHPPresentation) and [unit testing](http://phpoffice.github.io/PHPPresentation/coverage/develop/). You can learn more about PHPPresentation by reading the [Developers' Documentation](http://phppowerpoint.readthedocs.org/) and the [API Documentation](http://phpoffice.github.io/PHPPresentation/docs/develop/).

Read more about PHPPresentation:

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Getting started](#getting-started)
- [Contributing](#contributing)
- [Developers' Documentation](http://phppowerpoint.readthedocs.org/)
- [API Documentation](http://phpoffice.github.io/PHPPresentation/docs/master/)

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

- PHP 5.3+
- [Zip extension](http://php.net/manual/en/book.zip.php)
- [XML Parser extension](http://www.php.net/manual/en/xml.installation.php)
- [XMLWriter extension](http://php.net/manual/en/book.xmlwriter.php) (optional, used to write DOCX and ODT)

### Installation

It is recommended that you install the PHPPresentation library [through composer](http://getcomposer.org/). To do so, add
the following lines to your ``composer.json``.

```json
{
    "require": {
       "phpoffice/phppresentation": "dev-master"
    }
}
```

Alternatively, you can download the latest release from the [releases page](https://github.com/PHPOffice/PHPPresentation/releases).
In this case, you will have to register the autoloader. Register autoloading is required only if you do not use composer in your project.

```php
require_once 'path/to/PhpPresentation/src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();
```

## Getting started

The following is a basic usage example of the PHPPresentation library.

```php
// with your own install
require_once 'src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();

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

More examples are provided in the [samples folder](samples/). You can also read the [Developers' Documentation](http://phppowerpoint.readthedocs.org/) and the [API Documentation](http://phpoffice.github.io/PHPPresentation/docs/master/) for more detail.


## Contributing

We welcome everyone to contribute to PHPPresentation. Below are some of the things that you can do to contribute:

- Read [our contributing guide](https://github.com/PHPOffice/PHPPresentation/blob/master/CONTRIBUTING.md)
- [Fork us](https://github.com/PHPOffice/PHPPresentation/fork) and [request a pull](https://github.com/PHPOffice/PHPPresentation/pulls) to the [develop](https://github.com/PHPOffice/PHPPresentation/tree/develop) branch
- Submit [bug reports or feature requests](https://github.com/PHPOffice/PHPPresentation/issues) to GitHub
- Follow [@PHPOffice](https://twitter.com/PHPOffice) on Twitter
