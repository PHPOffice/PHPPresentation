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

use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Series;

/**
 * Test class for Pie3D element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D
 */
class Pie3DTest extends \PHPUnit_Framework_TestCase
{
    public function testData()
    {
        $object = new Pie3D();

        $this->assertInternalType('array', $object->getSeries());
        $this->assertEmpty($object->getSeries());

        $array = array(
            new Series(),
            new Series(),
        );

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->setSeries());
        $this->assertEmpty($object->getSeries());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->setSeries($array));
        $this->assertCount(count($array), $object->getSeries());
    }

    public function testSeries()
    {
        $object = new Pie3D();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getSeries());
    }

    public function testExplosion()
    {
        $value = rand(0, 100);
        
        $object = new Pie3D();

        $this->assertEquals(0, $object->getExplosion());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->setExplosion($value));
        $this->assertEquals($value, $object->getExplosion());
    }

    public function testHashCode()
    {
        $oSeries = new Series();

        $object = new Pie3D();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode().get_class($object)), $object->getHashCode());
    }
}
