![PHPPowerPoint](https://github.com/PHPOffice/PHPPowerPoint/raw/master/Documentation/assets/PHPPowerPointLogo.png "PHPPowerPoint")

# Developer Documentation

***Version 0.2.0***

PHPPowerPoint is a library written in pure PHP that provides a set of classes to write to different presentation file formats, i.e. OpenXML (.pptx) and OpenDocument (.odp). PHPPowerPoint is an open source project licensed under LGPL.

## Contents

- [Overview](#overview)
    - [Features](#features)
    - [Requirements](#requirements)
    - [Installation](#installation)
- [Objects](#objects)
    - [Document properties](#document-properties)
    - [Slides](#slides)
    - [Shapes](#shapes)
        - [Rich text](#rich-text)
        - [Line](#line)
        - [Chart](#chart)
        - [Drawing](#drawing)
        - [Table](#table)
    - [Styles](#styles)
        - [Fill](#fill)
        - [Border](#border)
        - [Alignment](#alignment)
        - [Font](#font)
        - [Bullet](#bullet)
        - [Color](#color)
- [Writers](#writers)
    - [OOXML](#ooxml)
    - [OpenDocument](#opendocument)
- [References](#references)

## Overview

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

## Objects

### Document properties

Use `getProperties` method to set document properties or metadata like author name and presentation title like below.

```php
$phpPowerPoint = new PHPPowerPoint();
$properties = $phpPowerPoint->getProperties();
$properties->setTitle('My presentation');
```

Available methods for setting document properties are:

- `setCategory`
- `setCompany`
- `setCreated`
- `setCreator`
- `setDescription`
- `setKeywords`
- `setLastModifiedBy`
- `setModified`
- `setSubject`
- `setTitle`

### Slides

Slides are pages in a presentation. Slides are stored as a zero based array in `PHPPowerPoint` object. Use `createSlide` to create a new slide and retrieve the slide for other operation such as creating shapes for that slide.

```php
$slide = $phpPowerPoint->createSlide();
$shape = $slide->createRichTextShape();
```

### Shapes

Shapes are objects that can be added to a slide. There are five types of shapes that can be used, i.e. [rich text](#rich-text), [line](#line), [chart](#chart), [drawing](#drawing), and [table](#table). Read the corresponding section of this manual for detail information of each shape.

Every shapes have common properties that you can set by using fluent interface.

- `width` in pixels
- `height` in pixels
- `offsetX` in pixels
- `offsetY` in pixels
- `rotation` in degrees
- `fill` see *[Fill](#fill)*
- `border` see *[Border](#border)*
- `shadow` see *[Shadow](#shadow)*
- `hyperlink`

Example:

```php
$richtext = $slide->createRichTextShape()
    ->setHeight(300)
    ->setWidth(600)
    ->setOffsetX(170)
    ->setOffsetY(180);
```

#### Rich text

Rich text shapes contain paragraphs of texts. To create a rich text shape, use `createRichTextShape` method of slide.

```php
$richtext = $slide->createRichTextShape()
```

Below are the properties that you can set for a rich text shape.

- `wrap`
- `autoFit`
- `horizontalOverflow`
- `verticalOverflow`
- `upright`
- `vertical`
- `columns`
- `bottomInset` in pixels
- `leftInset` in pixels
- `rightInset` in pixels
- `topInset` in pixels

Example:

```php
$richtext = $slide->createRichTextShape()
    ->setWrap(PHPPowerPoint_Shape_RichText::WRAP_SQUARE)
    ->setBottomInset(600);
```

Properties that can be set for each paragraphs are as follow.

- `alignment` see *[Alignment](#alignment)*
- `font` see *[Font](#font)*
- `bulletStyle` see *[Bullet](#bullet)*

#### Line

To create a line, use `createLineShape` method of slide.

```php
$line = $slide->createLineShape($fromX, $fromY, $toX, $toY);
```

#### Chart

To create a chart, use `createChartShape` method of slide.

```php
$chart = $slide->createChartShape();
```

#### Drawing

To create a drawing, use `createDrawingShape` method of slide.

```php
$drawing = $slide->createDrawingShape();
$drawing->setName('Unique name')
    ->setDescription('Description of the drawing')
    ->setPath('/path/to/drawing.filename');
```

#### Table

To create a table, use `createTableShape` method of slide.

```php
$table = $slide->createTableShape($columns);
```

### Styles

#### Fill

Use this style to define fill of a shape as example below.

```php
$shape->getFill()
    ->setFillType(PHPPowerPoint_Style_Fill::FILL_GRADIENT_LINEAR)
    ->setRotation(270)
    ->setStartColor(new PHPPowerPoint_Style_Color('FFCCCCCC'))
    ->setEndColor(new PHPPowerPoint_Style_Color('FFFFFFFF'));
```

Properties:

- `fillType`
- `rotation`
- `startColor`
- `endColor`

#### Border

Use this style to define border of a shape as example below.

```php
$shape->getBorder()
    ->setLineStyle(PHPPowerPoint_Style_Border::LINE_SINGLE)
    ->setLineWidth(4)
    ->getColor()->setARGB('FFC00000');
```

Properties:

- `lineWidth`
- `lineStyle`
- `dashStyle`
- `color`

#### Shadow

Use this style to define shadow of a shape as example below.

```php
$shape->getShadow()
    ->setVisible(true)
    ->setDirection(45)
    ->setDistance(10);
```

Properties:

- `visible`
- `blurRadius`
- `distance`
- `direction`
- `alignment`
- `color`
- `alpha`

#### Alignment

- `horizontal`
- `vertical`
- `level`
- `indent`
- `marginLeft`
- `marginRight`

#### Font

- `name`
- `bold`
- `italic`
- `superScript`
- `subScript`
- `underline`
- `strikethrough`
- `color`

#### Bullet

- `bulletType`
- `bulletFont`
- `bulletChar`
- `bulletNumericStyle`
- `bulletNumericStartAt`

#### Color

Colors can be applied to different objects, e.g. font or border.

```php
$textRun = $shape->createTextRun('Text');
$textRun->getFont()->setColor(new PHPPowerPoint_Style_Color('C00000'));
```

## Writers

Use the `IOFactory` object to write resulting presentation to a file.

```php
$writer = PHPPowerPoint_IOFactory::createWriter($phpPowerPoint, $writerName);
$writer->save('/path/to/result.document');
```

### OOXML

Use `PowerPoint2007` for `$writerName`.

```php
$writer = PHPPowerPoint_IOFactory::createWriter($phpPowerPoint, 'PowerPoint2007');
$writer->save('/path/to/result.pptx');
```

### OpenDocument

Use `ODPresentation` for `$writerName`.

```php
$writer = PHPPowerPoint_IOFactory::createWriter($phpPowerPoint, 'ODPresentation');
$writer->save('/path/to/result.odp');
```

## References

* [ISO/IEC 29500-1 Office Open XML File Formats](http://standards.iso.org/ittf/PubliclyAvailableStandards/c061750_ISO_IEC_29500-1_2012.zip) (2012)
* [Open Document Format for Office Applications (OpenDocument) Version 1.2](https://www.oasis-open.org/standards#opendocumentv1.2) (2011)
* [Ecma TC45 Office Open XML File Formats Standard](http://www.ecma-international.org/news/TC45_current_work/TC45_available_docs.htm) (2006)
* [OpenXML Explained](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2007/08/13/1970.aspx) (2007)
* [OASIS OpenDocument Essentials](http://books.evc-cit.info/odbook/book.html) (2005)
