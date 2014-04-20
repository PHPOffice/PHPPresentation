# PHPPowerPoint Manual

PHPPowerPoint is a library written in pure PHP that provides a set of classes to write to and read from different presentation file formats, i.e. OpenXML (.pptx) and OpenDocument (.odp). PHPPowerPoint is an open source project licensed under LGPL.

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

## Usages

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

Shapes are objects that can be added to a slide. There are five types of shapes that can be used, i.e. rich text, line, chart, drawing, and table.

To create a shape, use the following methods of a slide.

- `createRichTextShape`
- `createLineShape`
- `createChartShape`
- `createDrawingShape`
- `createTableShape`

Example:

```php
$richText = $slide->createRichTextShape();
$line = $slide->createLineShape();
```

Every shapes have common properties that you can set by using fluent interface.

- `width` in pixels
- `height` in pixels
- `offsetX` in pixels
- `offsetY` in pixels
- `rotation` in degrees
- `fill`
- `border`
- `shadow`
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

Rich text shapes contain paragraphs of texts. Below are the properties that you can set for a rich text shape.

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

- `alignment`
- `font`
- `bulletStyle`

#### Line

#### Chart

#### Drawing

#### Table

### Styles

#### Fill

- `fillType`
- `rotation`
- `startColor`
- `endColor`

#### Border

- `lineWidth`
- `lineStyle`
- `dashStyle`
- `color`

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
