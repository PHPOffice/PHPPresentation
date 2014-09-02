# ![PHPPowerPoint](https://github.com/PHPOffice/PHPPowerPoint/raw/master/docs/images/PHPPowerPointLogo.png "PHPPowerPoint")

[![Latest Stable Version](https://poser.pugx.org/phpoffice/phppowerpoint/v/stable.png)](https://packagist.org/packages/phpoffice/phppowerpoint)
[![Build Status](https://travis-ci.org/PHPOffice/PHPPowerPoint.svg?branch=master)](https://travis-ci.org/PHPOffice/PHPPowerPoint)
[![Code Quality](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/badges/quality-score.png?s=b5997ce59ac2816b4514f3a38de9900f6d492c1d)](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/)
[![Code Coverage](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/badges/coverage.png?s=742a98745725c562955440edc8d2c39d7ff5ae25)](https://scrutinizer-ci.com/g/PHPOffice/PHPPowerPoint/)
[![Total Downloads](https://poser.pugx.org/phpoffice/phppowerpoint/downloads.png)](https://packagist.org/packages/phpoffice/phppowerpoint)
[![License](https://poser.pugx.org/phpoffice/phppowerpoint/license.png)](https://packagist.org/packages/phpoffice/phppowerpoint)


PHPPowerPoint is a library written in pure PHP that provides a set of classes to write to different presentation file formats, i.e. Microsoft [Office Open XML](http://en.wikipedia.org/wiki/Office_Open_XML) (OOXML or OpenXML) or OASIS [Open Document Format for Office Applications](http://en.wikipedia.org/wiki/OpenDocument) (OpenDocument or ODF). 

PHPPowerPoint is an open source project licensed under the terms of [LGPL version 3](https://github.com/PHPOffice/PHPPowerPoint/blob/develop/COPYING.LESSER). PHPPowerPoint is aimed to be a high quality software product by incorporating [continuous integration](https://travis-ci.org/PHPOffice/PHPPowerPoint) and [unit testing](http://phpoffice.github.io/PHPPowerPoint/coverage/develop/). You can learn more about PHPPowerPoint by reading the [Developers' Documentation](http://phppowerpoint.readthedocs.org/) and the [API Documentation](http://phpoffice.github.io/PHPPowerPoint/docs/develop/).

Read more about PHPPowerPoint:

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Getting started](#getting-started)
- [Known issues](#known-issues)
- [Contributing](#contributing)
- [Developers' Documentation](http://phppowerpoint.readthedocs.org/)
- [API Documentation](http://phpoffice.github.io/PHPPowerPoint/docs/master/)

### Features

- Create an in-memory presentation representation
- Set presentation meta data (author, title, description, etc)
- Add slides from scratch or from existing one
- Supports different fonts and font styles
- Supports different formatting, styles, fills, gradients
- Supports hyperlinks and rich-text strings
- Add images with different styles (positioning, rotation, shadow)
- Set printing options (header, footer, page margins, paper size, orientation)
- Output to different file formats: PowerPoint 2007 (.pptx), OpenDocument Presentation (.odp), Serialized Presentation)
- ... and lots of other things!

### Requirements

PHPPowerPoint requires the following:

- PHP 5.3+
- [Zip extension](http://php.net/manual/en/book.zip.php)
- [XML Parser extension](http://www.php.net/manual/en/xml.installation.php)
- [XMLWriter extension](http://php.net/manual/en/book.xmlwriter.php) (optional, used to write DOCX and ODT)

### Installation

It is recommended that you install the PHPPowerPoint library [through composer](http://getcomposer.org/). To do so, add
the following lines to your ``composer.json``.

```json
{
    "require": {
       "phpoffice/phppowerpoint": "dev-master"
    }
}
```

Alternatively, you can download the latest release from the [releases page](https://github.com/PHPOffice/PHPPowerPoint/releases).
In this case, you will have to register the autoloader. Register autoloading is required only if you do not use composer in your project.

```php
require_once 'path/to/PhpPowerpoint/src/PhpPowerpoint/Autoloader.php';
\PhpOffice\PhpPowerpoint\Autoloader::register();
```

## Getting started

The following is a basic usage example of the PHPPowerPoint library.

```php
require_once 'src/PhpPowerpoint/Autoloader.php';
\PhpOffice\PhpPowerpoint\Autoloader::register();

$objPHPPowerPoint = new PhpPowerpoint();

// Create slide
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Create a shape (drawing)
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPowerPoint logo')
      ->setDescription('PHPPowerPoint logo')
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
$textRun = $shape->createTextRun('Thank you for using PHPPowerPoint!');
$textRun->getFont()->setBold(true)
                   ->setSize(60)
                   ->setColor( new Color( 'FFE06B20' ) );
                   
$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
$oWriterODP = IOFactory::createWriter($objPHPPowerPoint, 'ODPresentation');
$oWriterODP->save(__DIR__ . "/sample.odp");
```

More examples are provided in the [samples folder](samples/). You can also read the [Developers' Documentation](http://phppowerpoint.readthedocs.org/) and the [API Documentation](http://phpoffice.github.io/PHPPowerPoint/docs/master/) for more detail.


## Contributing

We welcome everyone to contribute to PHPPowerPoint. Below are some of the things that you can do to contribute:

- Read [our contributing guide](https://github.com/PHPOffice/PHPPowerPoint/blob/master/CONTRIBUTING.md)
- [Fork us](https://github.com/PHPOffice/PHPPowerPoint/fork) and [request a pull](https://github.com/PHPOffice/PHPPowerPoint/pulls) to the [develop](https://github.com/PHPOffice/PHPPowerPoint/tree/develop) branch
- Submit [bug reports or feature requests](https://github.com/PHPOffice/PHPPowerPoint/issues) to GitHub
- Follow [@PHPOffice](https://twitter.com/PHPOffice) on Twitter
