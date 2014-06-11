<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\DocumentLayout;
use PhpOffice\PhpPowerpoint\DocumentProperties;

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\PhpPowerpoint
 */
class PhpPowerpointTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new PhpPowerpoint();
        $slide = $object->getSlide();

        $this->assertEquals(new DocumentProperties(), $object->getProperties());
        $this->assertEquals(new DocumentLayout(), $object->getLayout());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->getSlide());
        $this->assertEquals(1, count($object->getAllSlides()));
        $this->assertEquals(0, $object->getIndex($slide));
        $this->assertEquals(1, $object->getSlideCount());
        $this->assertEquals(0, $object->getActiveSlideIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Iterator', $object->getSlideIterator());
    }

    /**
     * Test add external slide
     */
    public function testAddExternalSlide()
    {
        $origin = new PhpPowerpoint();
        $slide = $origin->getSlide();
        $object = new PhpPowerpoint();
        $object->addExternalSlide($slide);

        $this->assertEquals(2, $object->getSlideCount());
    }

    /**
     * Test copy presentation
     */
    public function testCopy()
    {
        $object = new PhpPowerpoint();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\PhpPowerpoint', $object->copy());
    }

    /**
     * Test remove slide by index exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage Slide index is out of bounds.
     */
    public function testRemoveSlideByIndexException()
    {
        $object = new PhpPowerpoint();
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
        $object = new PhpPowerpoint();
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
        $object = new PhpPowerpoint();
        $object->setActiveSlideIndex(1);
    }
}
