# Styles

## Fill

Use this style to define fill of a shape as example below.

``` php
<?php

$shape->getFill()
    ->setFillType(Fill::FILL_GRADIENT_LINEAR)
    ->setRotation(270)
    ->setStartColor(new Color('FFCCCCCC'))
    ->setEndColor(new Color('FFFFFFFF'));
```

Properties:

- `fillType`
- `rotation`
- `startColor`
- `endColor`

## Border

Use this style to define border of a shape as example below.

``` php
<?php

	$shape->getBorder()
		->setLineStyle(Border::LINE_SINGLE)
		->setLineWidth(4)
		->getColor()->setARGB('FFC00000');
```

Properties:

- `lineWidth`
- `lineStyle`
- `dashStyle`
- `color`

## Shadow

Use this style to define shadow of a shape as example below.

``` php
<?php

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

## Alignment

- `horizontal`
- `vertical`
- `level`
- `indent`
- `marginLeft`
- `marginRight`

### RTL / LTR

You can define if the alignment is RTL or LTR.

``` php
<?php

use PhpOffice\PhpPresentation\Style\Alignment;

$alignment = new Alignment();

// Set alignment to RTL
$alignment->setIsRTL(true);
// Set alignment to LTR
$alignment->setIsRTL(false);
// Is the alignment RTL?
echo $alignment->isRTL();
```

## Font

- `name`
- `bold`
- `italic`
- `superScript`
- `subScript`
- `underline`
- `strikethrough`
- `color`

### Format

Some formats are available : 

* `Font::FORMAT_LATIN`
* `Font::FORMAT_EAST_ASIAN`
* `Font::FORMAT_COMPLEX_SCRIPT`

``` php
<?php

use PhpOffice\PhpPresentation\Style\Font;

$font = new Font();

// Set format of font
$font->setFormat(Font::FORMAT_EAST_ASIAN);
// Get format of font
echo $font->getFormat();
```
## Bullet

- `bulletType`
- `bulletFont`
- `bulletChar`
- `bulletNumericStyle`
- `bulletNumericStartAt`

## Color

Colors can be applied to different objects, e.g. font or border.

``` php
<?php

$textRun = $shape->createTextRun('Text');
$textRun->getFont()->setColor(new Color('C00000'));
```
