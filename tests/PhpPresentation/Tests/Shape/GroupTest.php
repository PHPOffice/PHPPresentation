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

use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Group element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Group
 */
class GroupTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Group();

        self::assertEquals(0, $object->getOffsetX());
        self::assertEquals(0, $object->getOffsetY());
        self::assertEquals(0, $object->getExtentX());
        self::assertEquals(0, $object->getExtentY());
        self::assertCount(0, $object->getShapeCollection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Group', $object->setWidth(mt_rand(1, 100)));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Group', $object->setHeight(mt_rand(1, 100)));
    }

    public function testAdd(): void
    {
        $object = new Group();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AutoShape', $object->createAutoShape());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart', $object->createChartShape());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\File', $object->createDrawingShape());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Line', $object->createLineShape(10, 10, 10, 10));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $object->createRichTextShape());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table', $object->createTableShape());
        self::assertCount(6, $object->getShapeCollection());
    }

    public function testExtentX(): void
    {
        $object = new Group();
        $line1 = new Line(10, 20, 30, 50);
        $object->addShape($line1);

        self::assertEquals(20, $object->getExtentX());
    }

    public function testExtentY(): void
    {
        $object = new Group();
        $line1 = new Line(10, 20, 30, 50);
        $object->addShape($line1);

        self::assertEquals(30, $object->getExtentY());
    }

    public function testOffsetX(): void
    {
        $object = new Group();
        $line1 = new Line(10, 20, 30, 50);
        $object->addShape($line1);

        self::assertEquals($line1->getOffsetX(), $object->getOffsetX());

        self::assertInstanceOf(Group::class, $object->setOffsetX(mt_rand(1, 100)));
        self::assertEquals($line1->getOffsetX(), $object->getOffsetX());
    }

    public function testOffsetY(): void
    {
        $object = new Group();
        $line1 = new Line(10, 20, 30, 50);
        $object->addShape($line1);

        self::assertEquals($line1->getOffsetY(), $object->getOffsetY());

        self::assertInstanceOf(Group::class, $object->setOffsetY(mt_rand(1, 100)));
        self::assertEquals($line1->getOffsetY(), $object->getOffsetY());
    }

    public function testExtentsAndOffsetsForOneShape(): void
    {
        // We record initial values here because
        // PhpOffice\PhpPresentation\Shape\Line subtracts the offsets
        // from the extents to produce a raw width and height.
        $offsetX = 100;
        $offsetY = 100;
        $endX = 1000;
        $endY = 450;
        $extentX = $endX - $offsetX;
        $extentY = $endY - $offsetY;

        $object = new Group();
        $line1 = new Line($offsetX, $offsetY, $endX, $endY);
        $object->addShape($line1);

        self::assertEquals($offsetX, $object->getOffsetX());
        self::assertEquals($offsetY, $object->getOffsetY());
        self::assertEquals($extentX, $object->getExtentX());
        self::assertEquals($extentY, $object->getExtentY());
    }

    public function testExtentsAndOffsetsForTwoShapes(): void
    {
        // Since Groups and Slides cache offsets and extents on first
        // calculation, this test is separate from the above.
        // Should the calculation be performed every GET, this test can be
        // combined with the above.
        $offsetX = 100;
        $offsetY = 100;
        $endX = 1000;
        $endY = 450;
        $increase = 50;
        $extentX = ($endX - $offsetX) + $increase;
        $extentY = ($endY - $offsetY) + $increase;

        $line1 = new Line($offsetX, $offsetY, $endX, $endY);
        $line2 = new Line(
            $offsetX + $increase,
            $offsetY + $increase,
            $endX + $increase,
            $endY + $increase
        );

        $object = new Group();

        $object->addShape($line1);
        $object->addShape($line2);

        self::assertEquals($offsetX, $object->getOffsetX());
        self::assertEquals($offsetY, $object->getOffsetY());
        self::assertEquals($extentX, $object->getExtentX());
        self::assertEquals($extentY, $object->getExtentY());
    }
}
