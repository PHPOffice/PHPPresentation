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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Bar element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Bar
 */
class BarTest extends TestCase
{
    public function testData(): void
    {
        $object = new Bar();

        self::assertIsArray($object->getSeries());
        self::assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setSeries());
        self::assertEmpty($object->getSeries());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setSeries($array));
        self::assertCount(count($array), $object->getSeries());
    }

    public function testSeries(): void
    {
        $object = new Bar();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->addSeries(new Series()));
        self::assertCount(1, $object->getSeries());
    }

    public function testBarDirection(): void
    {
        $object = new Bar();
        self::assertEquals(Bar::DIRECTION_VERTICAL, $object->getBarDirection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarDirection(Bar::DIRECTION_HORIZONTAL));
        self::assertEquals(Bar::DIRECTION_HORIZONTAL, $object->getBarDirection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarDirection(Bar::DIRECTION_VERTICAL));
        self::assertEquals(Bar::DIRECTION_VERTICAL, $object->getBarDirection());
    }

    public function testBarGrouping(): void
    {
        $object = new Bar();
        self::assertEquals(Bar::GROUPING_CLUSTERED, $object->getBarGrouping());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarGrouping(Bar::GROUPING_CLUSTERED));
        self::assertEquals(Bar::GROUPING_CLUSTERED, $object->getBarGrouping());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarGrouping(Bar::GROUPING_STACKED));
        self::assertEquals(Bar::GROUPING_STACKED, $object->getBarGrouping());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarGrouping(Bar::GROUPING_PERCENTSTACKED));
        self::assertEquals(Bar::GROUPING_PERCENTSTACKED, $object->getBarGrouping());
    }

    public function testGapWidthPercent(): void
    {
        $value = mt_rand(0, 500);
        $object = new Bar();
        self::assertEquals(150, $object->getGapWidthPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent($value));
        self::assertEquals($value, $object->getGapWidthPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent(-1));
        self::assertEquals(0, $object->getGapWidthPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent(501));
        self::assertEquals(500, $object->getGapWidthPercent());
    }

    public function testOverlapWidthPercentDefaults(): void
    {
        $object = new Bar();
        self::assertEquals(0, $object->getOverlapWidthPercent());

        $object->setBarGrouping(Bar::GROUPING_STACKED);
        self::assertEquals(100, $object->getOverlapWidthPercent());
        $object->setBarGrouping(Bar::GROUPING_CLUSTERED);
        self::assertEquals(0, $object->getOverlapWidthPercent());
        $object->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);
        self::assertEquals(100, $object->getOverlapWidthPercent());
    }

    public function testOverlapWidthPercent(): void
    {
        $value = mt_rand(-100, 100);
        $object = new Bar();
        self::assertEquals(0, $object->getOverlapWidthPercent());
        self::assertInstanceOf(Bar::class, $object->setOverlapWidthPercent($value));
        self::assertEquals($value, $object->getOverlapWidthPercent());
        self::assertInstanceOf(Bar::class, $object->setOverlapWidthPercent(101));
        self::assertEquals(100, $object->getOverlapWidthPercent());
        self::assertInstanceOf(Bar::class, $object->setOverlapWidthPercent(-101));
        self::assertEquals(-100, $object->getOverlapWidthPercent());
    }

    public function testHashCode(): void
    {
        $oSeries = new Series();

        $object = new Bar();
        $object->addSeries($oSeries);

        self::assertEquals(md5($oSeries->getHashCode() . get_class($object)), $object->getHashCode());
    }
}
