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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Exception\UndefinedChartTypeException;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\PlotArea;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PlotArea element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\PlotArea
 */
class PlotAreaTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new PlotArea();

        self::assertInstanceOf(Axis::class, $object->getAxisX());
        self::assertInstanceOf(Axis::class, $object->getAxisY());
    }

    public function testHashIndex(): void
    {
        $object = new PlotArea();
        $value = mt_rand(1, 100);

        self::assertEmpty($object->getHashIndex());
        self::assertInstanceOf(PlotArea::class, $object->setHashIndex($value));
        self::assertEquals($value, $object->getHashIndex());
    }

    public function testHeight(): void
    {
        $object = new PlotArea();
        $value = mt_rand(0, 100);

        self::assertInstanceOf(PlotArea::class, $object->setHeight());
        self::assertEquals(0, $object->getHeight());
        self::assertInstanceOf(PlotArea::class, $object->setHeight($value));
        self::assertEquals($value, $object->getHeight());
    }

    public function testOffsetX(): void
    {
        $object = new PlotArea();
        $value = mt_rand(0, 100);

        self::assertInstanceOf(PlotArea::class, $object->setOffsetX());
        self::assertEquals(0, $object->getOffsetX());
        self::assertInstanceOf(PlotArea::class, $object->setOffsetX($value));
        self::assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY(): void
    {
        $object = new PlotArea();
        $value = mt_rand(0, 100);

        self::assertInstanceOf(PlotArea::class, $object->setOffsetY());
        self::assertEquals(0, $object->getOffsetY());
        self::assertInstanceOf(PlotArea::class, $object->setOffsetY($value));
        self::assertEquals($value, $object->getOffsetY());
    }

    public function testType(): void
    {
        $object = new PlotArea();

        self::assertInstanceOf(PlotArea::class, $object->setType(new Bar3D()));
        self::assertInstanceOf(AbstractType::class, $object->getType());
    }

    public function testTypeException(): void
    {
        $this->expectException(UndefinedChartTypeException::class);
        $this->expectExceptionMessage('The chart type has not been defined');

        $object = new PlotArea();
        $object->getType();
    }

    public function testWidth(): void
    {
        $object = new PlotArea();
        $value = mt_rand(0, 100);

        self::assertInstanceOf(PlotArea::class, $object->setWidth());
        self::assertEquals(0, $object->getWidth());
        self::assertInstanceOf(PlotArea::class, $object->setWidth($value));
        self::assertEquals($value, $object->getWidth());
    }
}
