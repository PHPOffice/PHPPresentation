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
 * @copyright   2009-2015 PHPPresentation contributors
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

        $this->assertIsArray($object->getSeries());
        $this->assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setSeries());
        $this->assertEmpty($object->getSeries());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setSeries($array));
        $this->assertCount(count($array), $object->getSeries());
    }

    public function testSeries(): void
    {
        $object = new Bar();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getSeries());
    }

    public function testBarDirection(): void
    {
        $object = new Bar();
        $this->assertEquals(Bar::DIRECTION_VERTICAL, $object->getBarDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarDirection(Bar::DIRECTION_HORIZONTAL));
        $this->assertEquals(Bar::DIRECTION_HORIZONTAL, $object->getBarDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarDirection(Bar::DIRECTION_VERTICAL));
        $this->assertEquals(Bar::DIRECTION_VERTICAL, $object->getBarDirection());
    }

    public function testBarGrouping(): void
    {
        $object = new Bar();
        $this->assertEquals(Bar::GROUPING_CLUSTERED, $object->getBarGrouping());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarGrouping(Bar::GROUPING_CLUSTERED));
        $this->assertEquals(Bar::GROUPING_CLUSTERED, $object->getBarGrouping());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarGrouping(Bar::GROUPING_STACKED));
        $this->assertEquals(Bar::GROUPING_STACKED, $object->getBarGrouping());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarGrouping(Bar::GROUPING_PERCENTSTACKED));
        $this->assertEquals(Bar::GROUPING_PERCENTSTACKED, $object->getBarGrouping());
    }

    public function testGapWidthPercent(): void
    {
        $value = mt_rand(0, 500);
        $object = new Bar();
        $this->assertEquals(150, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent($value));
        $this->assertEquals($value, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent(-1));
        $this->assertEquals(0, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent(501));
        $this->assertEquals(500, $object->getGapWidthPercent());
    }

    public function testOverlapWidthPercentDefaults(): void
    {
        $object = new Bar();
        $this->assertEquals(0, $object->getOverlapWidthPercent());

        $object->setBarGrouping(Bar::GROUPING_STACKED);
        $this->assertEquals(100, $object->getOverlapWidthPercent());
        $object->setBarGrouping(Bar::GROUPING_CLUSTERED);
        $this->assertEquals(0, $object->getOverlapWidthPercent());
        $object->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);
        $this->assertEquals(100, $object->getOverlapWidthPercent());
    }

    public function testOverlapWidthPercent(): void
    {
        $value = mt_rand(-100, 100);
        $object = new Bar();
        $this->assertEquals(0, $object->getOverlapWidthPercent());
        $this->assertInstanceOf(Bar::class, $object->setOverlapWidthPercent($value));
        $this->assertEquals($value, $object->getOverlapWidthPercent());
        $this->assertInstanceOf(Bar::class, $object->setOverlapWidthPercent(101));
        $this->assertEquals(100, $object->getOverlapWidthPercent());
        $this->assertInstanceOf(Bar::class, $object->setOverlapWidthPercent(-101));
        $this->assertEquals(-100, $object->getOverlapWidthPercent());
    }

    public function testHashCode(): void
    {
        $oSeries = new Series();

        $object = new Bar();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode() . get_class($object)), $object->getHashCode());
    }
}
