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

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\Exception\ShapeContainerAlreadyAssignedException;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Autoloader.
 */
class AbstractShapeTest extends TestCase
{
    /**
     * Register.
     */
    public function testConstruct(): void
    {
        $object = new RichText();

        self::assertEquals(0, $object->getOffsetX());
        self::assertEquals(0, $object->getOffsetY());
        self::assertEquals(0, $object->getHeight());
        self::assertEquals(0, $object->getRotation());
        self::assertEquals(0, $object->getWidth());
        self::assertInstanceOf(Border::class, $object->getBorder());
        self::assertEquals(Border::LINE_NONE, $object->getBorder()->getLineStyle());
        self::assertInstanceOf(Fill::class, $object->getFill());
        self::assertInstanceOf(Shadow::class, $object->getShadow());
    }

    public function testFill(): void
    {
        $object = new RichText();

        self::assertInstanceOf(AbstractShape::class, $object->setFill());
        self::assertNull($object->getFill());
        self::assertInstanceOf(AbstractShape::class, $object->setFill(new Fill()));
        self::assertInstanceOf(Fill::class, $object->getFill());
    }

    public function testHeight(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        self::assertInstanceOf(AbstractShape::class, $object->setHeight());
        self::assertEquals(0, $object->getHeight());
        self::assertInstanceOf(AbstractShape::class, $object->setHeight($value));
        self::assertEquals($value, $object->getHeight());
    }

    public function testHyperlink(): void
    {
        $object = new RichText();

        self::assertInstanceOf(AbstractShape::class, $object->setHyperlink());
        self::assertFalse($object->hasHyperlink());
        self::assertInstanceOf(Hyperlink::class, $object->getHyperlink());
        self::assertTrue($object->hasHyperlink());
        self::assertInstanceOf(AbstractShape::class, $object->setHyperlink(new Hyperlink('http://www.google.fr')));
        self::assertTrue($object->hasHyperlink());
        self::assertInstanceOf(Hyperlink::class, $object->getHyperlink());
        self::assertTrue($object->hasHyperlink());
    }

    public function testOffsetX(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        self::assertInstanceOf(AbstractShape::class, $object->setOffsetX());
        self::assertEquals(0, $object->getOffsetX());
        self::assertInstanceOf(AbstractShape::class, $object->setOffsetX($value));
        self::assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        self::assertInstanceOf(AbstractShape::class, $object->setOffsetY());
        self::assertEquals(0, $object->getOffsetY());
        self::assertInstanceOf(AbstractShape::class, $object->setOffsetY($value));
        self::assertEquals($value, $object->getOffsetY());
    }

    public function testRotation(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        self::assertInstanceOf(AbstractShape::class, $object->setRotation());
        self::assertEquals(0, $object->getRotation());
        self::assertInstanceOf(AbstractShape::class, $object->setRotation($value));
        self::assertEquals($value, $object->getRotation());
    }

    public function testShadow(): void
    {
        $object = new RichText();

        self::assertInstanceOf(AbstractShape::class, $object->setShadow());
        self::assertNull($object->getShadow());
        self::assertInstanceOf(AbstractShape::class, $object->setShadow(new Shadow()));
        self::assertInstanceOf(Shadow::class, $object->getShadow());
    }

    public function testWidth(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        self::assertInstanceOf(AbstractShape::class, $object->setWidth());
        self::assertEquals(0, $object->getWidth());
        self::assertInstanceOf(AbstractShape::class, $object->setWidth($value));
        self::assertEquals($value, $object->getWidth());
    }

    public function testWidthAndHeight(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        self::assertInstanceOf(AbstractShape::class, $object->setWidthAndHeight());
        self::assertEquals(0, $object->getWidth());
        self::assertEquals(0, $object->getHeight());
        self::assertInstanceOf(AbstractShape::class, $object->setWidthAndHeight($value));
        self::assertEquals($value, $object->getWidth());
        self::assertEquals(0, $object->getHeight());
        self::assertInstanceOf(AbstractShape::class, $object->setWidthAndHeight($value, $value));
        self::assertEquals($value, $object->getWidth());
        self::assertEquals($value, $object->getHeight());
    }

    public function testPlaceholder(): void
    {
        $object = new RichText();
        self::assertFalse($object->isPlaceholder(), 'Standard Shape should not be a placeholder object');
        self::assertNull($object->getPlaceholder());
        self::assertInstanceOf(
            AbstractShape::class,
            $object->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_TITLE))
        );
        self::assertTrue($object->isPlaceholder());
        self::assertInstanceOf(Placeholder::class, $object->getPlaceholder());
        self::assertEquals('title', $object->getPlaceholder()->getType());

        $object = new RichText();
        self::assertFalse($object->isPlaceholder(), 'Standard Shape should not be a placeholder object');
        $placeholder = new Placeholder(Placeholder::PH_TYPE_TITLE);
        $placeholder->setType(Placeholder::PH_TYPE_SUBTITLE);
        self::assertInstanceOf(AbstractShape::class, $object->setPlaceHolder($placeholder));
        self::assertTrue($object->isPlaceholder());
        self::assertInstanceOf(Placeholder::class, $object->getPlaceholder());
        self::assertEquals('subTitle', $object->getPlaceholder()->getType());
    }

    public function testContainer(): void
    {
        $object = new RichText();
        $object2 = new RichText();
        $object2->setWrap(RichText::WRAP_NONE);
        $oSlide = new Slide();
        $oSlide->addShape($object2);

        self::assertNull($object->getContainer());
        self::assertInstanceOf(AbstractShape::class, $object->setContainer($oSlide));
        self::assertInstanceOf(Slide::class, $object->getContainer());
        self::assertInstanceOf(AbstractShape::class, $object->setContainer(null, true));
        self::assertNull($object->getContainer());
    }

    public function testContainerException(): void
    {
        $object = new RichText();
        $oSlide = new Slide();

        self::assertNull($object->getContainer());
        self::assertInstanceOf(AbstractShape::class, $object->setContainer($oSlide));
        self::assertInstanceOf(Slide::class, $object->getContainer());
        $this->expectException(ShapeContainerAlreadyAssignedException::class);
        $this->expectExceptionMessage('The shape PhpOffice\PhpPresentation\AbstractShape has already a container assigned');
        $object->setContainer(null);
    }
}
