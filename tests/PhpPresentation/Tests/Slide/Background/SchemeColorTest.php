<?php

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\SchemeColor;
use PhpOffice\PhpPresentation\Style\SchemeColor as StyleSchemeColor;
use PHPUnit\Framework\TestCase;

class SchemeColorTest extends TestCase
{
    public function testBasic()
    {
        $oStyle = new StyleSchemeColor();

        $object = new SchemeColor();

        static::assertNull($object->getSchemeColor());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\SchemeColor', $object->setSchemeColor($oStyle));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\SchemeColor', $object->getSchemeColor());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\SchemeColor', $object->setSchemeColor());
        static::assertNull($object->getSchemeColor());
    }
}
