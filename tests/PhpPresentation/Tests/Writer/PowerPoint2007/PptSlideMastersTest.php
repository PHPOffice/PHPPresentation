<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use ArrayObject;
use DOMDocument;
use DOMElement;
use DOMXPath;
use PhpOffice\PhpPresentation\Shape\Drawing\File as ShapeDrawingFile;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\PptSlideMasters;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PowerPoint2007.
 *
 * @coversDefaultClass \PowerPoint2007
 */
class PptSlideMastersTest extends TestCase
{
    public function testWriteSlideMasterRelationships(): void
    {
        $writer = new PptSlideMasters();
        /** @var \PHPUnit\Framework\MockObject\MockObject|SlideMaster $slideMaster */
        $slideMaster = $this->getMockBuilder(SlideMaster::class)
            ->setMethods(['getAllSlideLayouts', 'getRelsIndex', 'getShapeCollection'])
            ->getMock();

        $layouts = [new SlideLayout($slideMaster)];

        $slideMaster->expects($this->once())
            ->method('getAllSlideLayouts')
            ->will($this->returnValue($layouts));

        $collection = new ArrayObject();
        $collection[] = new ShapeDrawingFile();
        $collection[] = new ShapeDrawingFile();
        $collection[] = new ShapeDrawingFile();

        $slideMaster->expects($this->exactly(2))
            ->method('getShapeCollection')
            ->will($this->returnValue($collection));

        $data = $writer->writeSlideMasterRelationships($slideMaster);

        $dom = new DOMDocument();
        $dom->loadXml($data);

        $xpath = new DOMXPath($dom);
        $xpath->registerNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $list = $xpath->query('//r:Relationship');

        $this->assertEquals(5, $list->length);

        foreach (range(0, 4) as $id) {
            /** @var DOMElement $domItem */
            $domItem = $list->item($id);
            $this->assertInstanceOf(DOMElement::class, $domItem);
            $this->assertEquals('rId' . (string) ($id + 1), $domItem->getAttribute('Id'));
        }
    }
}
