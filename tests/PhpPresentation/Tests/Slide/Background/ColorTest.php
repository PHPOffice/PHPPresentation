<?php

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\Color;
use PhpOffice\PhpPresentation\Style\Color as StyleColor;
use PHPUnit\Framework\TestCase;

class ColorTest extends TestCase
{
    public function testColor()
    {
        $object = new Color();

        $oStyleColor = new StyleColor();
        $oStyleColor->setRGB('123456');

        static::assertNull($object->getColor());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->setColor($oStyleColor));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->setColor());
        static::assertNull($object->getColor());
    }
}
