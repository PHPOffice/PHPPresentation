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

use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\File;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Slide\Animation;
use PhpOffice\PhpPresentation\Slide\Transition;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class SlideTest extends TestCase
{
    public function testExtents(): void
    {
        $object = new Slide();
        self::assertSame(0, $object->getExtentX());

        $object = new Slide();
        self::assertSame(0, $object->getExtentY());
    }

    public function testOffset(): void
    {
        $object = new Slide();
        self::assertSame(0, $object->getOffsetX());

        $object = new Slide();
        self::assertSame(0, $object->getOffsetY());
    }

    public function testParent(): void
    {
        $object = new Slide();
        self::assertNull($object->getParent());

        $oPhpPresentation = new PhpPresentation();
        $object = new Slide($oPhpPresentation);
        self::assertInstanceOf(PhpPresentation::class, $object->getParent());
    }

    public function testSlideMasterId(): void
    {
        $value = mt_rand(1, 100);
        $object = new Slide();
        self::assertEquals(1, $object->getSlideMasterId());
        self::assertInstanceOf(Slide::class, $object->setSlideMasterId());
        self::assertEquals(1, $object->getSlideMasterId());
        self::assertInstanceOf(Slide::class, $object->setSlideMasterId($value));
        self::assertEquals($value, $object->getSlideMasterId());
    }

    public function testAnimations(): void
    {
        $oStub = new class() extends Animation {
        };

        $object = new Slide();
        self::assertIsArray($object->getAnimations());
        self::assertCount(0, $object->getAnimations());
        self::assertInstanceOf(Slide::class, $object->addAnimation($oStub));
        self::assertIsArray($object->getAnimations());
        self::assertCount(1, $object->getAnimations());
        self::assertInstanceOf(Slide::class, $object->setAnimations());
        self::assertIsArray($object->getAnimations());
        self::assertCount(0, $object->getAnimations());
        self::assertInstanceOf(Slide::class, $object->setAnimations([$oStub]));
        self::assertIsArray($object->getAnimations());
        self::assertCount(1, $object->getAnimations());
    }

    public function testBackground(): void
    {
        $oStub = new class() extends AbstractBackground {
        };

        $object = new Slide();
        self::assertNull($object->getBackground());
        self::assertInstanceOf(Slide::class, $object->setBackground($oStub));
        self::assertInstanceOf(AbstractBackground::class, $object->getBackground());
        self::assertInstanceOf(Slide::class, $object->setBackground());
        self::assertNull($object->getBackground());
    }

    public function testGroup(): void
    {
        $object = new Slide();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Group', $object->createGroup());
    }

    public function testName(): void
    {
        $object = new Slide();
        self::assertNull($object->getName());
        self::assertInstanceOf(Slide::class, $object->setName('AAAA'));
        self::assertEquals('AAAA', $object->getName());
        self::assertInstanceOf(Slide::class, $object->setName());
        self::assertNull($object->getName());
    }

    public function testTransition(): void
    {
        $object = new Slide();
        $oTransition = new Transition();
        self::assertNull($object->getTransition());
        self::assertInstanceOf(Slide::class, $object->setTransition());
        self::assertNull($object->getTransition());
        self::assertInstanceOf(Slide::class, $object->setTransition($oTransition));
        self::assertInstanceOf(Transition::class, $object->getTransition());
        self::assertInstanceOf(Slide::class, $object->setTransition(null));
        self::assertNull($object->getTransition());
    }

    public function testVisible(): void
    {
        $object = new Slide();
        self::assertTrue($object->isVisible());
        self::assertInstanceOf(Slide::class, $object->setIsVisible(false));
        self::assertFalse($object->isVisible());
        self::assertInstanceOf(Slide::class, $object->setIsVisible());
        self::assertTrue($object->isVisible());
    }

    public function testAddShape(): void
    {
        $slide = new Slide();
        self::assertInstanceOf(ShapeContainerInterface::class, $slide);
        $shape = new File();

        self::assertCount(0, $slide->getShapeCollection());

        $slide->addShape($shape);
        self::assertInstanceOf(File::class, $shape);
        self::assertEquals($slide, $shape->getContainer());
        self::assertInstanceOf(Slide::class, $shape->getContainer());

        self::assertCount(1, $slide->getShapeCollection());
        self::assertEquals([$shape], $slide->getShapeCollection());
    }

    public function testMoveShape(): void
    {
        $slide = new Slide();
        $first = $slide->createRichTextShape();
        $second = $slide->createRichTextShape();
        $third = $slide->createRichTextShape();

        self::assertInstanceOf(Slide::class, $slide->moveShape($first, 2));
        self::assertSame([$second, $third, $first], $slide->getShapeCollection());

        $slide->moveShape($first, 0);
        self::assertSame([$first, $second, $third], $slide->getShapeCollection());
    }

    public function testMoveShapeAfterAHoleInTheKeys(): void
    {
        $slide = new Slide();
        $slide->createRichTextShape();
        $second = $slide->createRichTextShape();
        $third = $slide->createRichTextShape();

        // the key of a shape is no longer the place it sits in
        $slide->unsetShape(0);
        $slide->moveShape($second, 1);

        self::assertSame([$third, $second], $slide->getShapeCollection());
    }

    public function testMoveShapeNotInTheContainer(): void
    {
        $slide = new Slide();
        $slide->createRichTextShape();

        $this->expectException(InvalidParameterException::class);
        $slide->moveShape((new Slide())->createRichTextShape(), 0);
    }

    public function testMoveShapeOutOfBounds(): void
    {
        $slide = new Slide();
        $shape = $slide->createRichTextShape();

        $this->expectException(OutOfBoundsException::class);
        $slide->moveShape($shape, 1);
    }

    public function testCreateDrawingShape(): void
    {
        $slide = new Slide();

        self::assertCount(0, $slide->getShapeCollection());

        $shape = $slide->createDrawingShape();
        self::assertInstanceOf(File::class, $shape);
        self::assertEquals($slide, $shape->getContainer());
        self::assertInstanceOf(Slide::class, $shape->getContainer());

        self::assertCount(1, $slide->getShapeCollection());
        self::assertEquals([$shape], $slide->getShapeCollection());
    }
}
