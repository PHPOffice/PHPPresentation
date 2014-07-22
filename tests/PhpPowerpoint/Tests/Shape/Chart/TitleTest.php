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

use PhpOffice\PhpPowerpoint\Shape\Chart\Title;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Font;

/**
 * Test class for Title element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Chart\Title
 */
class TitleTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Title();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Alignment', $object->getAlignment());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Font', $object->getFont());
        $this->assertEquals('Calibri', $object->getFont()->getName());
        $this->assertEquals(18, $object->getFont()->getSize());
    }

    public function testAlignment()
    {
        $object = new Title();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setAlignment(new Alignment()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Alignment', $object->getAlignment());
    }

    public function testFont()
    {
        $object = new Title();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setFont(new Font()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Font', $object->getFont());
    }

    public function testHashIndex()
    {
        $object = new Title();
        $value = rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testHeight()
    {
        $object = new Title();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setHeight());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }

    public function testOffsetX()
    {
        $object = new Title();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setOffsetX());
        $this->assertEquals(0.01, $object->getOffsetX());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setOffsetX($value));
        $this->assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY()
    {
        $object = new Title();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setOffsetY());
        $this->assertEquals(0.01, $object->getOffsetY());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setOffsetY($value));
        $this->assertEquals($value, $object->getOffsetY());
    }

    public function testText()
    {
        $object = new Title();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setText());
        $this->assertNull($object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setText('AAAA'));
        $this->assertEquals('AAAA', $object->getText());
    }

    public function testVisible()
    {
        $object = new Title();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setVisible());
        $this->assertTrue($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setVisible(true));
        $this->assertTrue($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setVisible(false));
        $this->assertFalse($object->isVisible());
    }

    public function testWidth()
    {
        $object = new Title();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }
}
