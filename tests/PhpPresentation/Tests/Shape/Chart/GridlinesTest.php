<?php

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Style\Outline;
use PHPUnit\Framework\TestCase;

class GridlinesTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Gridlines();

        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testGetSetOutline(): void
    {
        $object = new Gridlines();

        /** @var Outline $oStub */
        $oStub = $this->getMockBuilder(Outline::class)->getMock();

        $this->assertInstanceOf(Outline::class, $object->getOutline());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Gridlines', $object->setOutline($oStub));
        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }
}
