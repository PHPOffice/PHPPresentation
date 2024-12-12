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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart\Type;

use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Bar3D element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D
 */
class Bar3DTest extends TestCase
{
    public function testData(): void
    {
        $object = new Bar3D();

        self::assertIsArray($object->getSeries());
        self::assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setSeries());
        self::assertEmpty($object->getSeries());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setSeries($array));
        self::assertCount(count($array), $object->getSeries());
    }

    public function testSeries(): void
    {
        $object = new Bar3D();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->addSeries(new Series()));
        self::assertCount(1, $object->getSeries());
    }

    public function testBarDirection(): void
    {
        $object = new Bar3D();
        self::assertEquals(Bar3D::DIRECTION_VERTICAL, $object->getBarDirection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarDirection(Bar3D::DIRECTION_HORIZONTAL));
        self::assertEquals(Bar3D::DIRECTION_HORIZONTAL, $object->getBarDirection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarDirection(Bar3D::DIRECTION_VERTICAL));
        self::assertEquals(Bar3D::DIRECTION_VERTICAL, $object->getBarDirection());
    }

    public function testBarGrouping(): void
    {
        $object = new Bar3D();
        self::assertEquals(Bar3D::GROUPING_CLUSTERED, $object->getBarGrouping());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarGrouping(Bar3D::GROUPING_CLUSTERED));
        self::assertEquals(Bar3D::GROUPING_CLUSTERED, $object->getBarGrouping());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarGrouping(Bar3D::GROUPING_STACKED));
        self::assertEquals(Bar3D::GROUPING_STACKED, $object->getBarGrouping());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarGrouping(Bar3D::GROUPING_PERCENTSTACKED));
        self::assertEquals(Bar3D::GROUPING_PERCENTSTACKED, $object->getBarGrouping());
    }

    public function testGapWidthPercent(): void
    {
        $value = mt_rand(0, 500);
        $object = new Bar3D();
        self::assertEquals(150, $object->getGapWidthPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setGapWidthPercent($value));
        self::assertEquals($value, $object->getGapWidthPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setGapWidthPercent(-1));
        self::assertEquals(0, $object->getGapWidthPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setGapWidthPercent(501));
        self::assertEquals(500, $object->getGapWidthPercent());
    }

    public function testHashCode(): void
    {
        $oSeries = new Series();

        $object = new Bar3D();
        $object->addSeries($oSeries);

        self::assertEquals(md5($oSeries->getHashCode() . get_class($object)), $object->getHashCode());
    }
}
