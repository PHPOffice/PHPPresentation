<?php
/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\PhpPresentation;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class PhpPresentationTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new PhpPresentation();
        $slide = $object->getSlide();

        $this->assertEquals(new DocumentProperties(), $object->getProperties());
        $this->assertEquals(new DocumentLayout(), $object->getLayout());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->getSlide());
        $this->assertEquals(1, count($object->getAllSlides()));
        $this->assertEquals(0, $object->getIndex($slide));
        $this->assertEquals(1, $object->getSlideCount());
        $this->assertEquals(0, $object->getActiveSlideIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Iterator', $object->getSlideIterator());
    }

    /**
     * Test add external slide
     */
    public function testAddExternalSlide()
    {
        $origin = new PhpPresentation();
        $slide = $origin->getSlide();
        $object = new PhpPresentation();
        $object->addExternalSlide($slide);

        $this->assertEquals(2, $object->getSlideCount());
    }

    /**
     * Test copy presentation
     */
    public function testCopy()
    {
        $object = new PhpPresentation();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->copy());
    }

    /**
     * Test remove slide by index exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage Slide index is out of bounds.
     */
    public function testRemoveSlideByIndexException()
    {
        $object = new PhpPresentation();
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
        $object = new PhpPresentation();
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
        $object = new PhpPresentation();
        $object->setActiveSlideIndex(1);
    }
}
