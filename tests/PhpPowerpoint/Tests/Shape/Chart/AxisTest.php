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

use PhpOffice\PhpPowerpoint\Shape\Chart\Legend;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Style\Font;

/**
 * Test class for Legend element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Chart\Legend
 */
class AxisTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Legend();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Font', $object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getBorder());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Fill', $object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Alignment', $object->getAlignment());
    }

    public function testAlignment()
    {
        $object = new Legend();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setAlignment(new Alignment()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Alignment', $object->getAlignment());
    }

    public function testFont()
    {
        $object = new Legend();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setFont(new Font()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Font', $object->getFont());
    }

    public function testHashIndex()
    {
        $object = new Legend();
        $value = rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testHeight()
    {
        $object = new Legend();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setHeight());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }

    public function testOffsetX()
    {
        $object = new Legend();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setOffsetX());
        $this->assertEquals(0, $object->getOffsetX());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setOffsetX($value));
        $this->assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY()
    {
        $object = new Legend();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setOffsetY());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setOffsetY($value));
        $this->assertEquals($value, $object->getOffsetY());
    }

    public function testPosition()
    {
        $object = new Legend();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setPosition());
        $this->assertEquals(Legend::POSITION_RIGHT, $object->getPosition());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setPosition(Legend::POSITION_BOTTOM));
        $this->assertEquals(Legend::POSITION_BOTTOM, $object->getPosition());
    }

    public function testVisible()
    {
        $object = new Legend();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setVisible());
        $this->assertTrue($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setVisible(true));
        $this->assertTrue($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setVisible(false));
        $this->assertFalse($object->isVisible());
    }

    public function testWidth()
    {
        $object = new Legend();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }
}
