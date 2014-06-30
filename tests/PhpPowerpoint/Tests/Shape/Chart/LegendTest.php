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

namespace PhpOffice\PhpPowerpoint\Tests\Shape\Chart;

use PhpOffice\PhpPowerpoint\Shape\Chart\Axis;

/**
 * Test class for Axis element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Chart\Axis
 */
class LegendTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Axis();

        $this->assertEquals('Axis Title', $object->getTitle());
    }

    public function testFormatCode()
    {
        $object = new Axis();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Axis', $object->setFormatCode());
        $this->assertEquals('', $object->getFormatCode());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Axis', $object->setFormatCode('AAAA'));
        $this->assertEquals('AAAA', $object->getFormatCode());
    }

    public function testHashIndex()
    {
        $object = new Axis();
        $value = rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Axis', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testTitle()
    {
        $object = new Axis();
        $this->assertEquals('Axis Title', $object->getTitle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Axis', $object->setTitle('AAAA'));
        $this->assertEquals('AAAA', $object->getTitle());
    }
}
