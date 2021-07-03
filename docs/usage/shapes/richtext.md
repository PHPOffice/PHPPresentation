# RichText

Rich text shapes contain paragraphs of texts. To create a rich text shape, use `createRichTextShape` method of slide.

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

Properties that can be set for each paragraphs are as follow.

- `alignment` <!-- see *[Alignment](#alignment)*-->
- `bulletStyle` see *[Bullet](#bullet)*
- `lineSpacing` see *[LineSpacing](#linespacing)*
- `font` <!-- see *[Font](#font)*-->

## Bullet

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

## LineSpacing

For a paragraph, you can define the line spacing.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;

$paragraph = new Paragraph();
$paragraph->setLineSpacing(200);
$lineSpacing = $paragraph->getLineSpacing();
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