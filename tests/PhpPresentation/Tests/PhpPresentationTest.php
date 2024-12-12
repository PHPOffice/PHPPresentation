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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Slide;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class PhpPresentationTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new PhpPresentation();
        $slide = $object->getSlide();

        self::assertEquals(new DocumentProperties(), $object->getDocumentProperties());
        self::assertEquals(new DocumentLayout(), $object->getLayout());
        self::assertInstanceOf(Slide::class, $object->getSlide());
        self::assertCount(1, $object->getAllSlides());
        self::assertEquals(0, $object->getIndex($slide));
        self::assertEquals(1, $object->getSlideCount());
        self::assertEquals(0, $object->getActiveSlideIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Iterator', $object->getSlideIterator());
    }

    public function testProperties(): void
    {
        $object = new PhpPresentation();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getDocumentProperties());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->setDocumentProperties(new DocumentProperties()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->getDocumentProperties());
    }

    public function testPresentationProperties(): void
    {
        $object = new PhpPresentation();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->getPresentationProperties());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->setPresentationProperties(new PresentationProperties()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->getPresentationProperties());
    }

    /**
     * Test add external slide.
     */
    public function testAddExternalSlide(): void
    {
        $origin = new PhpPresentation();
        $slide = $origin->getSlide();
        $object = new PhpPresentation();
        $object->addExternalSlide($slide);

        self::assertEquals(2, $object->getSlideCount());
    }

    /**
     * Test copy presentation.
     */
    public function testCopy(): void
    {
        $object = new PhpPresentation();
        $object->createSlide();

        $copy = $object->copy();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $copy);
        self::assertEquals(2, $copy->getSlideCount());
    }

    /**
     * Test remove slide by index exception.
     */
    public function testRemoveSlideByIndexException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1) is out of bounds (0, 0)');

        $object = new PhpPresentation();
        $object->removeSlideByIndex(1);
    }

    /**
     * Test get slide exception.
     */
    public function testGetSlideException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1) is out of bounds (0, 0)');

        $object = new PhpPresentation();
        $object->getSlide(1);
    }

    /**
     * Test set active slide index exception.
     */
    public function testSetActiveSlideIndexException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1) is out of bounds (0, 0)');

        $object = new PhpPresentation();
        $object->setActiveSlideIndex(1);
    }

    public function testAddSlideAtStart(): void
    {
        $presentation = new PhpPresentation();
        $presentation->removeSlideByIndex(0);
        $slide1 = new Slide($presentation);
        $slide1->setName('Slide 1');
        $slide2 = new Slide($presentation);
        $slide2->setName('Slide 2');
        $slide3 = new Slide($presentation);
        $slide3->setName('Slide 3');

        $presentation->addSlide($slide1);
        $presentation->addSlide($slide2);
        $presentation->addSlide($slide3, 0);

        self::assertEquals('Slide 3', $presentation->getSlide(0)->getName());
        self::assertEquals('Slide 1', $presentation->getSlide(1)->getName());
        self::assertEquals('Slide 2', $presentation->getSlide(2)->getName());
    }

    public function testAddSlideAtMiddle(): void
    {
        $presentation = new PhpPresentation();
        $presentation->removeSlideByIndex(0);
        $slide1 = new Slide($presentation);
        $slide1->setName('Slide 1');
        $slide2 = new Slide($presentation);
        $slide2->setName('Slide 2');
        $slide3 = new Slide($presentation);
        $slide3->setName('Slide 3');

        $presentation->addSlide($slide1);
        $presentation->addSlide($slide2);
        $presentation->addSlide($slide3, 1);

        self::assertEquals('Slide 1', $presentation->getSlide(0)->getName());
        self::assertEquals('Slide 3', $presentation->getSlide(1)->getName());
        self::assertEquals('Slide 2', $presentation->getSlide(2)->getName());
    }

    public function testAddSlideAtEnd(): void
    {
        $presentation = new PhpPresentation();
        $presentation->removeSlideByIndex(0);
        $slide1 = new Slide($presentation);
        $slide1->setName('Slide 1');
        $slide2 = new Slide($presentation);
        $slide2->setName('Slide 2');
        $slide3 = new Slide($presentation);
        $slide3->setName('Slide 3');

        $presentation->addSlide($slide1);
        $presentation->addSlide($slide2);
        $presentation->addSlide($slide3);

        self::assertEquals('Slide 1', $presentation->getSlide(0)->getName());
        self::assertEquals('Slide 2', $presentation->getSlide(1)->getName());
        self::assertEquals('Slide 3', $presentation->getSlide(2)->getName());
    }
}
