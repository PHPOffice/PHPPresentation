<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\DocumentLayout;
use PhpOffice\PhpPowerpoint\DocumentProperties;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;

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
