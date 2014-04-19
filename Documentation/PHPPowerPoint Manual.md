# PHPPowerPoint

PHPPowerPoint is a library written in pure PHP that provides a set of classes to write to and read from different presentation file formats. PHPPowerPoint is an open source project licensed under LGPL.

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

```php
$richText = $slide->createRichTextShape();
$line = $slide->createLineShape();
```

Below are the methods of slide that can be used to create shapes.

- `createRichTextShape`
- `createLineShape`
- `createChartShape`
- `createDrawingShape`
- `createTableShape`

#### Rich text

#### Line

#### Chart

#### Drawing

#### Table

### Styles

- Alignment
- Border
- Bullet
- Color
- Fill
- Font
