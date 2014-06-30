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

use PhpOffice\PhpPowerpoint\Shape\RichText\Run;
use PhpOffice\PhpPowerpoint\Style\Font;

/**
 * Test class for Run element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\RichText\Run
 */
class RunTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $object = new Run();
        $this->assertEquals('', $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Font', $object->getFont());

        $object = new Run('BBB');
        $this->assertEquals('BBB', $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Font', $object->getFont());
    }

    public function testFont()
    {
        $object = new Run();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Run', $object->setFont(new Font()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Font', $object->getFont());
    }

    public function testText()
    {
        $object = new Run();
        $this->assertEquals('', $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Run', $object->setText());
        $this->assertEquals('', $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Run', $object->setText('AAA'));
        $this->assertEquals('AAA', $object->getText());

        $object = new Run('BBB');
        $this->assertEquals('BBB', $object->getText());
    }

    /**
     * Test get/set hash index
     */
    public function testHashCode()
    {
        $object = new Run();
        $this->assertEquals(md5($object->getFont()->getHashCode().get_class($object)), $object->getHashCode());
    }
}
