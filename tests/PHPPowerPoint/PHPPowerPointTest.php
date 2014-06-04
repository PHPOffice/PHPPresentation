<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

/**
 * Test class for PHPPowerPoint
 *
 * @coversDefaultClass PHPPowerPoint
 */
class PHPPowerPointTest extends PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new PHPPowerPoint();
        $slide = $object->getSlide();

        $this->assertEquals(new PHPPowerPoint_DocumentProperties(), $object->getProperties());
        $this->assertEquals(new PHPPowerPoint_DocumentLayout(), $object->getLayout());
        $this->assertInstanceOf('PHPPowerPoint_Slide', $object->getSlide());
        $this->assertEquals(1, count($object->getAllSlides()));
        $this->assertEquals(0, $object->getIndex($slide));
        $this->assertEquals(1, $object->getSlideCount());
        $this->assertEquals(0, $object->getActiveSlideIndex());
        $this->assertInstanceOf('PHPPowerPoint_Slide_Iterator', $object->getSlideIterator());
    }

    /**
     * Test add external slide
     */
    public function testAddExternalSlide()
    {
        $origin = new PHPPowerPoint();
        $slide = $origin->getSlide();
        $object = new PHPPowerPoint();
        $object->addExternalSlide($slide);

        $this->assertEquals(2, $object->getSlideCount());
    }

    /**
     * Test copy presentation
     */
    public function testCopy()
    {
        $object = new PHPPowerPoint();
        $this->assertInstanceOf('PHPPowerPoint', $object->copy());
    }

    /**
     * Test remove slide by index exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage Slide index is out of bounds.
     */
    public function testRemoveSlideByIndexException()
    {
        $object = new PHPPowerPoint();
        $object->removeSlideByIndex(1);
    }

    /**
     * Test get slide exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage Slide index is out of bounds.
     */
    public function testGetSlideException()
    {
        $object = new PHPPowerPoint();
        $object->getSlide(1);
    }

    /**
     * Test set active slide index exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage Active slide index is out of bounds.
     */
    public function testSetActiveSlideIndexException()
    {
        $object = new PHPPowerPoint();
        $object->setActiveSlideIndex(1);
    }
}
