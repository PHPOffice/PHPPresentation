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

use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Series;

/**
 * Test class for Bar3D element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D
 */
class Bar3DTest extends \PHPUnit_Framework_TestCase
{
    public function testData()
    {
        $object = new Bar3D();

        $this->assertInternalType('array', $object->getSeries());
        $this->assertEmpty($object->getSeries());

        $array = array(
            new Series(),
            new Series(),
        );

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setSeries());
        $this->assertEmpty($object->getSeries());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setSeries($array));
        $this->assertCount(count($array), $object->getSeries());
    }

    public function testSeries()
    {
        $object = new Bar3D();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getSeries());
    }
    
    public function testBarDirection()
    {
        $object = new Bar3D();
        $this->assertEquals(Bar3D::DIRECTION_VERTICAL, $object->getBarDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarDirection(Bar3D::DIRECTION_HORIZONTAL));
        $this->assertEquals(Bar3D::DIRECTION_HORIZONTAL, $object->getBarDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarDirection(Bar3D::DIRECTION_VERTICAL));
        $this->assertEquals(Bar3D::DIRECTION_VERTICAL, $object->getBarDirection());
    }

    public function testBarGrouping()
    {
        $object = new Bar3D();
        $this->assertEquals(Bar3D::GROUPING_CLUSTERED, $object->getBarGrouping());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarGrouping(Bar3D::GROUPING_CLUSTERED));
        $this->assertEquals(Bar3D::GROUPING_CLUSTERED, $object->getBarGrouping());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarGrouping(Bar3D::GROUPING_STACKED));
        $this->assertEquals(Bar3D::GROUPING_STACKED, $object->getBarGrouping());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setBarGrouping(Bar3D::GROUPING_PERCENTSTACKED));
        $this->assertEquals(Bar3D::GROUPING_PERCENTSTACKED, $object->getBarGrouping());
    }

    public function testGapWidthPercent()
    {
        $value = rand(0, 500);
        $object = new Bar3D();
        $this->assertEquals(150, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setGapWidthPercent($value));
        $this->assertEquals($value, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setGapWidthPercent(-1));
        $this->assertEquals(0, $object->getGapWidthPercent());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Bar3D', $object->setGapWidthPercent(501));
        $this->assertEquals(500, $object->getGapWidthPercent());
    }

    public function testHashCode()
    {
        $oSeries = new Series();

        $object = new Bar3D();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode().get_class($object)), $object->getHashCode());
    }
}
