# ![PHPPowerPoint](https://github.com/PHPOffice/PHPPowerPoint/raw/master/Documentation/assets/PHPPowerPointLogo.png "PHPPowerPoint")

[![Build Status](https://travis-ci.org/PHPOffice/PHPPowerPoint.svg?branch=master)](https://travis-ci.org/PHPOffice/PHPPowerPoint)


PHPPowerPoint is a library written in pure PHP that provides a set of classes to write to different presentation file formats, i.e. OpenXML (.pptx) and OpenDocument (.odp). PHPPowerPoint is an open source project licensed under [LGPL](LICENSE.md).

### Features

- Create an in-memory presentation representation
- Set presentation meta data (author, title, description, etc)
- Add slides from scratch or from existing one
- Supports different fonts and font styles
- Supports different formatting, styles, fills, gradients
- Supports hyperlinks and rich-text strings
- Add images with different styles (positioning, rotation, shadow)
- Set printing options (header, footer, page margins, paper size, orientation)
- Output to different file formats: PowerPoint 2007 (.pptx), OpenDocument Presentation (.odp), Serialized Spreadsheet)
- ... and lots of other things!

### Requirements

The following requirements should be met prior to using PHPPowerPoint:

- PHP version 5.2 or higher
- PHP extension php_zip enabled
- PHP extension php_xml enabled

### Installation

To install and use PHPPowerPoint, copy the contents of the `Classes` folder and include `PHPPowerPoint.php` somewhere in your code like below.

```php
include_once '/path/to/Classes/PHPPowerPoint.php';
```

After that, you can use the library by creating a new instance of the class.

```php
$phpPowerPoint = new PHPPowerPoint();
```

### Want to learn more?

[Read the manual](Documentation/PHPPowerPointDocumentation.md).

### Want to contribute?

[Fork us on GitHub](https://github.com/PHPOffice/PHPPowerPoint)!
