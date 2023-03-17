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
- `columns`
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

## Column Spacing

For a paragraph, you can define the column spacing.

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