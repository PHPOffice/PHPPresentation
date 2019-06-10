<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Drawing\File as ShapeDrawingFile;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\PptSlideMasters;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class PptSlideMastersTest extends TestCase
{
    public function testWriteSlideMasterRelationships()
    {
        $writer = new PptSlideMasters();
        /** @var \PHPUnit_Framework_MockObject_MockObject|SlideMaster $slideMaster */
        $slideMaster = $this->getMockBuilder('PhpOffice\\PhpPresentation\\Slide\\SlideMaster')
            ->setMethods(array('getAllSlideLayouts', 'getRelsIndex', 'getShapeCollection'))
            ->getMock();

        $layouts = array(new SlideLayout($slideMaster));

        $slideMaster->expects($this->once())
            ->method('getAllSlideLayouts')
            ->will($this->returnValue($layouts));

        $collection = new \ArrayObject();
        $collection[] = new ShapeDrawingFile();
        $collection[] = new ShapeDrawingFile();
        $collection[] = new ShapeDrawingFile();

        $slideMaster->expects($this->exactly(2))
            ->method('getShapeCollection')
            ->will($this->returnValue($collection));

        $data = $writer->writeSlideMasterRelationships($slideMaster);

        $dom = new \DomDocument();
        $dom->loadXml($data);

        $xpath = new \DomXpath($dom);
        $xpath->registerNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $list = $xpath->query('//r:Relationship');

        $this->assertEquals(5, $list->length);

        $this->assertEquals('rId1', $list->item(0)->getAttribute('Id'));
        $this->assertEquals('rId2', $list->item(1)->getAttribute('Id'));
        $this->assertEquals('rId3', $list->item(2)->getAttribute('Id'));
        $this->assertEquals('rId4', $list->item(3)->getAttribute('Id'));
        $this->assertEquals('rId5', $list->item(4)->getAttribute('Id'));
    }
}
