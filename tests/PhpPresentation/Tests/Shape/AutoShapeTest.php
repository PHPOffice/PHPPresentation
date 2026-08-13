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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Style\Outline;
use PHPUnit\Framework\MockObject\MockObject;
use PHPUnit\Framework\TestCase;

class AutoShapeTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new AutoShape();

        self::assertEquals(AutoShape::TYPE_HEART, $object->getType());
        self::assertEquals('', $object->getText());
        self::assertInstanceOf(Outline::class, $object->getOutline());
        self::assertNotEmpty($object->getHashCode());
    }

    public function testOutline(): void
    {
        /** @var MockObject&Outline $mock */
        $mock = $this->getMockBuilder(Outline::class)->getMock();

        $object = new AutoShape();
        self::assertInstanceOf(Outline::class, $object->getOutline());
        self::assertInstanceOf(AutoShape::class, $object->setOutline($mock));
        self::assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testText(): void
    {
        $object = new AutoShape();

        self::assertEquals('', $object->getText());
        self::assertInstanceOf(AutoShape::class, $object->setText('Text'));
        self::assertEquals('Text', $object->getText());
    }

    public function testType(): void
    {
        $object = new AutoShape();

        self::assertEquals(AutoShape::TYPE_HEART, $object->getType());
        self::assertInstanceOf(AutoShape::class, $object->setType(AutoShape::TYPE_HEXAGON));
        self::assertEquals(AutoShape::TYPE_HEXAGON, $object->getType());
    }

    public function testNoRadiusByDefaultIsNull(): void
    {
        $shape = new AutoShape();
        self::assertNull($shape->getRoundRectCorner());
    }

    public function testRoundRectCornerPixelRadiusStoredAndAffectsHash(): void
    {
        $width = 200;  // px
        $height = 100;  // px
        $px1 = 5;  // softer radius
        $px2 = 10; // larger radius

        $shape1 = (new AutoShape())
            ->setType(AutoShape::TYPE_ROUNDED_RECTANGLE)
            ->setWidth($width)
            ->setHeight($height)
            ->setRoundRectCorner($px1);

        $shape2 = (clone $shape1)->setRoundRectCorner($px2);

        self::assertSame($px1, $shape1->getRoundRectCorner());
        self::assertSame($px2, $shape2->getRoundRectCorner());

        // Hash must differ when radius differs
        self::assertNotSame($shape1->getHashCode(), $shape2->getHashCode());
    }
}
