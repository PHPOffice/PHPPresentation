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

namespace PhpOffice\PhpPowerpoint\Tests\Shape;

use PhpOffice\PhpPowerpoint\Shape\Line;
use PhpOffice\PhpPowerpoint\Style\Border;

/**
 * Test class for memory drawing element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Line
 */
class LineTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $value = rand(1, 100);
        $object = new Line($value, $value, $value, $value);

        $this->assertEquals(Border::LINE_SINGLE, $object->getBorder()->getLineStyle());
        $this->assertEquals($value, $object->getOffsetX());
        $this->assertEquals($value, $object->getOffsetY());
        $this->assertEquals(0, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInternalType('string', $object->getHashCode());
    }
}
