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

namespace PhpOffice\PhpPowerpoint\Tests\Shape\RichText;

use PhpOffice\PhpPowerpoint\Shape\RichText\BreakElement;

/**
 * Test class for BreakElement element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\RichText\BreakElement
 */
class BreakElementTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testText()
    {
        $object = new BreakElement();
        $this->assertEquals("\r\n", $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\BreakElement', $object->setText());
        $this->assertEquals("\r\n", $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\BreakElement', $object->setText('AAA'));
        $this->assertEquals("\r\n", $object->getText());
    }

    public function testFont()
    {
        $object = new BreakElement();
        $this->assertNull($object->getFont());
    }

    /**
     * Test get/set hash index
     */
    public function testHashCode()
    {
        $object = new BreakElement();
        $this->assertEquals(md5(get_class($object)), $object->getHashCode());
    }
}
