<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Writer\PowerPoint2007\PptSlideMasters;
use PhpOffice\PhpPresentation\Slide\SlideLayout;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class PptSlideMastersTest extends \PHPUnit_Framework_TestCase
{
    public function testWriteSlideMasterRelationships()
    {
        $writer = new PptSlideMasters();
        $slideMaster = $this->getMockBuilder('PhpOffice\\PhpPresentation\\Slide\\SlideMaster')
            ->setMethods(['getAllSlideLayouts', 'getRelsIndex', 'getShapeCollection'])
            ->getMock();

        $layouts = [new SlideLayout($slideMaster)];

        $slideMaster->expects($this->once())
            ->method('getAllSlideLayouts')
            ->will($this->returnValue($layouts));

        $slideMaster->expects($this->exactly(2))
            ->method('getShapeCollection')
            ->will($this->returnValue(new \ArrayObject()));

        $data = $writer->writeSlideMasterRelationships($slideMaster);
    }
}
