<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shape\Chart\Type;

use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Series;

/**
 * Test class for Bar3D element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D
 */
class Bar3DTest extends \PHPUnit_Framework_TestCase
{
    public function testData()
    {
        $object = new Bar3D();

        $this->assertInternalType('array', $object->getData());
        $this->assertEmpty($object->getData());

        $array = array(
            new Series(),
            new Series(),
        );

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Type\\Bar3D', $object->setData());
        $this->assertEmpty($object->getData());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Type\\Bar3D', $object->setData($array));
        $this->assertCount(count($array), $object->getData());
    }

    public function testSerties()
    {
        $object = new Bar3D();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Type\\Bar3D', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getData());
    }

    public function testHashCode()
    {
        $oSeries = new Series();

        $object = new Bar3D();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode().get_class($object)), $object->getHashCode());
    }
}
