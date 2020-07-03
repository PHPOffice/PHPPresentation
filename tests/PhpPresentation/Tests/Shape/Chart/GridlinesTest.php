<?php

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;

class GridlinesTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Gridlines();

        $this->assertInstanceOf('PhpOffice\PhpPresentation\Style\Outline', $object->getOutline());
    }

    public function testGetSetOutline()
    {
        $object = new Gridlines();

        $oStub = $this->getMockBuilder('PhpOffice\PhpPresentation\Style\Outline')->getMock();

        $this->assertInstanceOf('PhpOffice\PhpPresentation\Style\Outline', $object->getOutline());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Gridlines', $object->setOutline($oStub));
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Style\Outline', $object->getOutline());
    }
}
