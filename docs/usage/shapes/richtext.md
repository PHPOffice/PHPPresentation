# RichText

Rich text shapes contain paragraphs of texts. To create a rich text shape, use `createRichTextShape` method of slide.

Each rich text can contain multiples paragraphs.
Each paragraph can contain:
- a `TextElement`
- a `BreakElement`
- a `Run`

Below are the properties that you can set for a rich text shape.

- `wrap`
- `autoFit`
- `fontScale` : font scale (in percentage) when autoFit = `RichText::AUTOFIT_NORMAL`
- `lnSpcReduction` : line spacing reduction (in percentage) when autoFit = `RichText::AUTOFIT_NORMAL`
- `horizontalOverflow`
- `verticalOverflow`
- `upright`
- `vertical`
- `columns` see *Columns*
- `columnsRTL` see *Columns*
- `bottomInset` in pixels
- `leftInset` in pixels
- `rightInset` in pixels
- `topInset` in pixels
- `autoShrinkHorizontal` (boolean)
- `autoShrinkVertical` (boolean)
- `columnSpacing` see *Column Spacing*

Properties that can be set for each paragraphs are as follow.

- `alignment` <!-- see *[Alignment](#alignment)*-->
- `bulletStyle` see *[Bullet](#bullet)*
- `lineSpacing` see *Line Spacing*
- `font` <!-- see *[Font](#font)*-->

## Columns

For a rich text, you can lay the text out in more than one column. The spacing between them is
given in pixels.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText;

$richText = new RichText();
$richText->setColumns(3);
$columns = $richText->getColumns();
```

The columns are ordered left to right unless every paragraph of the shape reads right to left, in
which case they follow the text. `setColumnsRTL()` says the order outright and overrides that:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText;

$richText = new RichText();
$richText->setColumns(3);
// right to left columns, whatever direction the text runs in
$richText->setColumnsRTL(true);
// null, the default, takes the order from the paragraphs
$richText->setColumnsRTL();
```

`hasColumnsRTL()` gives back what was set, `null` included; `isColumnsRTL()` gives the order the
writers use, taking it from the paragraphs when nothing was set.

## Column Spacing

For a rich text, you can define the spacing between its columns, in pixels. It has no effect on a
shape left with the single column it starts with.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText;

$richText = new RichText();
$richText->setColumnSpacing(200);
$columnSpacing = $richText->getColumnSpacing();
```

## Hyperlink

For a rich text, you can define the hyperlink.

Example:

```php
<?php

use PhpOffice\PhpPresentation\Shape\RichText;

$richText = new RichText();
$richText->getHyperlink()->setUrl('https://phpoffice.github.io/PHPPresentation/');

```

### Link to another slide

A hyperlink can point at another slide of the same deck rather than at a URL.

Example:

```php
<?php

use PhpOffice\PhpPresentation\Shape\RichText;

$richText = new RichText();
$richText->getHyperlink()->setSlideNumber(3);

```

Each format addresses the slide in its own way. The PowerPoint2007 Writer writes a relationship to
the slide part; the ODPresentation Writer writes the name of the page, because that is how ODF
addresses a slide. Every page therefore carries a `draw:name`, and a slide with no name of its own
is named after its position, the way LibreOffice names the pages of a deck.

### Use of Text Color

!!! warning
    Available only on the PowerPoint2007 Reader/Writer

Hyperlinks can be set to use the text color instead of the default theme color.

Example:

```php
<?php

use PhpOffice\PhpPresentation\Shape\RichText;

$richText = new RichText();
$richText->getHyperlink()->setUrl('https://phpoffice.github.io/PHPPresentation/');
$richText->getHyperlink()->setIsTextColorUsed(true);

```

## Paragraph
### Bullet

For a paragraph, you can define the bullet style.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Bullet;

$paragraph = new Paragraph();
$paragraph->getBulletStyle();
```

With the bullet style, you can define the char, the font, the color and the type.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;

$paragraph = new Paragraph();
$paragraph->getBulletStyle()->setBulletChar('-');
$paragraph->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
$paragraph->getBulletStyle()->setBulletColor(new Color(Color::COLOR_RED));
```

A numbered list is the `Bullet::TYPE_NUMERIC` type. The scheme it is numbered with is one of the
`Bullet::NUMERIC_*` constants, and the number it starts at is `setBulletNumericStartAt()`.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Bullet;

$paragraph = new Paragraph();
$paragraph->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
$paragraph->getBulletStyle()->setBulletNumericStyle(Bullet::NUMERIC_ALPHALCPARENBOTH);
$paragraph->getBulletStyle()->setBulletNumericStartAt(3);
```

### Line Spacing

For a paragraph, you can define the line spacing.
By default, mode is in percent (`Paragraph::LINE_SPACING_MODE_PERCENT`), but you can use the point mode (`Paragraph::LINE_SPACING_MODE_POINT`).

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;

$paragraph = new Paragraph();
$paragraph->setLineSpacing(200);
$lineSpacing = $paragraph->getLineSpacing();

$paragraph->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_POINT);
$lineSpacingMode = $paragraph->getLineSpacingMode();
```

### Spacing

For a paragraph, you can define the spacing before and after the paragraph in point
Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;

$paragraph = new Paragraph();
$paragraph->setSpacingAfter(12);
$spacingAfter = $paragraph->getSpacingAfter();

$paragraph->setSpacingBefore(34);
$spacingBefore = $paragraph->getSpacingBefore();
```

## Run

For a run, you can define the language.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Run;

$run = new Run();
$run->setLanguage('fr-FR');
```
## Field

A field is a run whose text the application recomputes: the number of the slide it ended up on,
the number of slides there are, the date the presentation is opened on. It stands among the runs
of a paragraph rather than taking the whole shape, so a sentence can be part written and part
computed.

The text a field is created with is what it stands in for: an application that does not know the
kind of field shows that text instead of computing one.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Field;

$paragraph = $slide->createRichTextShape()->getActiveParagraph();
$paragraph->createTextRun('page ');
$paragraph->createField(Field::TYPE_SLIDENUM, '<nr.>');
$paragraph->createTextRun(' of ');
$paragraph->createField(Field::TYPE_SLIDECOUNT, '12');
```

`Field::TYPE_SLIDENUM`, `Field::TYPE_SLIDECOUNT` and `Field::TYPE_DATETIME` are the kinds named
here, but the kind is an open string rather than a list of constants, because that is what
`a:fld/@type` is. Any name the reading application knows can be given.

### The kinds an application knows

`slidenum`, `datetime` and `datetime1` to `datetime13` are the ones the format reserves. The rest
are LibreOffice's, which reads and writes them all.

| Kind | What it says |
|---|---|
| `slidenum` | the number of the slide |
| `slidecount` | the number of slides |
| `slidename` | the name of the slide |
| `author` | who the presentation belongs to |
| `file` | path and file name |
| `file1` / `file2` / `file3` | path / file name without extension / file name with extension |
| `datetime` | the date, short |
| `datetime1` | `13/02/1996` |
| `datetime2` | `Tuesday, February 13, 1996` |
| `datetime3` | `13 February 1996` |
| `datetime4` | `February 13, 1996` |
| `datetime5` | `13-Feb-96` |
| `datetime6` | `February 96` |
| `datetime7` | `Feb-96` |
| `datetime8` | `13/02/1996 13:49` |
| `datetime9` | `13/02/1996 13:49:38` |
| `datetime10` | `13:49` |
| `datetime11` | `13:49:38` |
| `datetime12` | `01:49 PM` |
| `datetime13` | `01:49:38 PM` |

An application that does not know a kind shows the text the field stands in for, so a kind it has
no answer for costs the liveness of that one field and nothing else.

### What each writer does with a field

The PowerPoint2007 writer writes every kind, as the `a:fld` it is.

OpenDocument names a field by what it is rather than by how it is formatted, so the ODPresentation
writer maps the kinds onto the elements it has -- `text:page-number`, `text:page-count`,
`text:date`, `text:time`, `text:author-name`, `text:file-name` -- and all fourteen dated formats
come down to a date or a time. `slidename` has no element in OpenDocument 1.2, so it is written as
the text it stands in for, and so is any other kind. The date format itself would travel as a data
style, which this writer does not write.

The PowerPoint97 writer writes what the field stands in for, as the text it is.
