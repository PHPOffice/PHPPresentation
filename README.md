# PHPPowerPoint - OpenXML - Read, Write and Create PowerPoint documents in PHP

Project providing a set of classes for the PHP programming language, which allow you to write to and read from different file formats, like PowerPoint 2007, ... This project is built around Microsoft's OpenXML standard and PHP.
Checkout the Features this class set provides, such as setting presentation meta data (author, title, description, ...), adding slides, adding images to your presentation and much, much more!

## Want to contribute?

Fork us!

## Requirements

The following requirements should be met prior to using PHPPowerPoint:

* PHP version 5.2 or higher
* PHP extension php_zip enabled
* PHP extension php_xml enabled

## Installation

Installation is quite easy: copy the contents of the Classes folder to any location
in your application required.

Afterwards, make sure you can include all PHPPowerPoint files. This can be achieved by
respecting a base folder structure, or by setting the PHP include path, for example:

```php
set_include_path(get_include_path() . PATH_SEPARATOR . '/path/to/PHPPowerPoint/');
```

## License

PHPPowerPoint is licensed under [LGPL (GNU LESSER GENERAL PUBLIC LICENSE)](LICENSE.md)