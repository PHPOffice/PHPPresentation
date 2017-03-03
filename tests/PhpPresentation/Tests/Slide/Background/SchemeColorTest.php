<?php

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\SchemeColor;
use PhpOffice\PhpPresentation\Style\SchemeColor as StyleSchemeColor;

class SchemeColorTest extends \PHPUnit_Framework_TestCase
{
    public function testBasic()
    {
        $oStyle = new StyleSchemeColor();

        $object = new SchemeColor();

        $this->assertNull($object->getSchemeColor());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\SchemeColor', $object->setSchemeColor($oStyle));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\SchemeColor', $object->getSchemeColor());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\SchemeColor', $object->setSchemeColor());
        $this->assertNull($object->getSchemeColor());
    }
}
