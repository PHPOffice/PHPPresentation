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
use PhpOffice\PhpPresentation\PresentationProperties;

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

        $this->assertEquals(new DocumentProperties(), $object->getDocumentProperties());
        $this->assertEquals(new DocumentLayout(), $object->getLayout());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->getSlide());
        $this->assertEquals(1, count($object->getAllSlides()));
        $this->assertEquals(0, $object->getIndex($slide));
        $this->assertEquals(1, $object->getSlideCount());
        $this->assertEquals(0, $object->getActiveSlideIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Iterator', $object->getSlideIterator());
    }

    public function testProperties()
    {
        $object = new PhpPresentation();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getProperties());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getDocumentProperties());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->setProperties(new DocumentProperties()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getProperties());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getDocumentProperties());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->setDocumentProperties(new DocumentProperties()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getProperties());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getDocumentProperties());
    }

    public function testPresentationProperties()
    {
        $object = new PhpPresentation();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->getPresentationProperties());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->setPresentationProperties(new PresentationProperties()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->getPresentationProperties());
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

    /**
     * @deprecated
     */
    public function testMarkAsFinal()
    {
        $object = new PhpPresentation();
        $this->assertFalse($object->isMarkedAsFinal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->markAsFinal(true));
        $this->assertTrue($object->isMarkedAsFinal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->markAsFinal(false));
        $this->assertFalse($object->isMarkedAsFinal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->markAsFinal());
        $this->assertTrue($object->isMarkedAsFinal());
    }

    /**
     * @deprecated
     */
    public function testZoom()
    {
        $object = new PhpPresentation();
        $this->assertEquals(1, $object->getZoom());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setZoom(0.3));
        $this->assertEquals(0.3, $object->getZoom());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setZoom());
        $this->assertEquals(1, $object->getZoom());
    }
}
