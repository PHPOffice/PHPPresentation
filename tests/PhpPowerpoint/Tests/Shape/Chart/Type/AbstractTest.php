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

use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter;

/**
 * Test class for Scatter element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter
 */
class AbstractTest extends \PHPUnit_Framework_TestCase
{
    public function testAxis()
    {
        $object = new Scatter();

        $this->assertTrue($object->hasAxisX());
        $this->assertTrue($object->hasAxisY());
    }

    public function testHashIndex()
    {
        $object = new Scatter();
        $value = rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Type\\Scatter', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }
}
