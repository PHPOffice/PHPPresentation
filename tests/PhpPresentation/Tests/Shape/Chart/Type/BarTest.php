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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart\Type;

use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Series;

/**
 * Test class for Bar element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\Type\Bar
 */
class BarTest extends \PHPUnit_Framework_TestCase
{
    public function testData()
    {
        $object = new Bar();

        $this->assertInternalType('array', $object->getSeries());
        $this->assertEmpty($object->getSeries());

        $array = array(
            new Series(),
            new Series(),
        );

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setSeries());
        $this->assertEmpty($object->getSeries());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setSeries($array));
        $this->assertCount(count($array), $object->getSeries());
    }

    public function testSeries()
    {
        $object = new Bar();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getSeries());
    }
    
    public function testBarDirection()
    {
        $object = new Bar();
        $this->assertEquals(Bar::DIRECTION_VERTICAL, $object->getBarDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarDirection(Bar::DIRECTION_HORIZONTAL));
        $this->assertEquals(Bar::DIRECTION_HORIZONTAL, $object->getBarDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setBarDirection(Bar::DIRECTION_VERTICAL));
        $this->assertEquals(Bar::DIRECTION_VERTICAL, $object->getBarDirection());
    }

    public function testBarGrouping()
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

    public function testGapWidthPercent()
    {
        $value = rand(0, 500);
        $object = new Bar();
        $this->assertEquals(150, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent($value));
        $this->assertEquals($value, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent(-1));
        $this->assertEquals(0, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar', $object->setGapWidthPercent(501));
        $this->assertEquals(500, $object->getGapWidthPercent());
    }

    public function testHashCode()
    {
        $oSeries = new Series();

        $object = new Bar();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode().get_class($object)), $object->getHashCode());
    }
}
