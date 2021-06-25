<?php

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\Color;
use PhpOffice\PhpPresentation\Style\Color as StyleColor;
use PHPUnit\Framework\TestCase;

class ColorTest extends TestCase
{
    public function testColor(): void
    {
        $object = new Color();

        $oStyleColor = new StyleColor();
        $oStyleColor->setRGB('123456');

        $this->assertNull($object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->setColor($oStyleColor));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->setColor());
        $this->assertNull($object->getColor());
    }
}
