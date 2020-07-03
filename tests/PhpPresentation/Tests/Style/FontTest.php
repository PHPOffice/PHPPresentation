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

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class FontTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Font();
        $this->assertEquals('Calibri', $object->getName());
        $this->assertEquals(10, $object->getSize());
        $this->assertFalse($object->isBold());
        $this->assertFalse($object->isItalic());
        $this->assertFalse($object->isSuperScript());
        $this->assertFalse($object->isSubScript());
        $this->assertFalse($object->isStrikethrough());
        $this->assertEquals(Font::UNDERLINE_NONE, $object->getUnderline());
        $this->assertEquals(0, $object->getCharacterSpacing());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
    }

    /**
     * Test get/set color
     * @expectedException \Exception
     * @expectedExceptionMessage $pValue must be an instance of \PhpOffice\PhpPresentation\Style\Color
     */
    public function testSetGetColorException()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setColor());
    }

    /**
     * Test get/set Character Spacing
     */
    public function testSetGetCharacterSpacing()
    {
        $object = new Font();
        $this->assertEquals(0, $object->getCharacterSpacing());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setCharacterSpacing(0));
        $this->assertEquals(0, $object->getCharacterSpacing());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setCharacterSpacing(10));
        $this->assertEquals(1000, $object->getCharacterSpacing());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setCharacterSpacing());
        $this->assertEquals(0, $object->getCharacterSpacing());
    }

    /**
     * Test get/set color
     */
    public function testSetGetColor()
    {
        $object = new Font();
        $this->assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setColor(new Color(Color::COLOR_BLUE)));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertEquals(Color::COLOR_BLUE, $object->getColor()->getARGB());
    }

    /**
     * Test get/set name
     */
    public function testSetGetName()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setName());
        $this->assertEquals('Calibri', $object->getName());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setName(''));
        $this->assertEquals('Calibri', $object->getName());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setName('Arial'));
        $this->assertEquals('Arial', $object->getName());
    }

    /**
     * Test get/set size
     */
    public function testSetGetSize()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSize());
        $this->assertEquals(10, $object->getSize());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSize(''));
        $this->assertEquals(10, $object->getSize());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSize($value));
        $this->assertEquals($value, $object->getSize());
    }

    /**
     * Test get/set underline
     */
    public function testSetGetUnderline()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setUnderline());
        $this->assertEquals(FONT::UNDERLINE_NONE, $object->getUnderline());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setUnderline(''));
        $this->assertEquals(FONT::UNDERLINE_NONE, $object->getUnderline());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setUnderline(FONT::UNDERLINE_DASH));
        $this->assertEquals(FONT::UNDERLINE_DASH, $object->getUnderline());
    }

    /**
     * Test get/set bold
     */
    public function testSetIsBold()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold());
        $this->assertFalse($object->isBold());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold(''));
        $this->assertFalse($object->isBold());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold(false));
        $this->assertFalse($object->isBold());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold(true));
        $this->assertTrue($object->isBold());
    }

    /**
     * Test get/set italic
     */
    public function testSetIsItalic()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic());
        $this->assertFalse($object->isItalic());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic(''));
        $this->assertFalse($object->isItalic());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic(false));
        $this->assertFalse($object->isItalic());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic(true));
        $this->assertTrue($object->isItalic());
    }

    /**
     * Test get/set strikethrough
     */
    public function testSetIsStriketrough()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough());
        $this->assertFalse($object->isStrikethrough());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough(''));
        $this->assertFalse($object->isStrikethrough());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough(false));
        $this->assertFalse($object->isStrikethrough());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough(true));
        $this->assertTrue($object->isStrikethrough());
    }

    /**
     * Test get/set subscript
     */
    public function testSetIsSubScript()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript());
        $this->assertFalse($object->isSubScript());
        $this->assertTrue($object->isSuperScript());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(''));
        $this->assertFalse($object->isSubScript());
        $this->assertTrue($object->isSuperScript());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(false));
        $this->assertFalse($object->isSubScript());
        $this->assertTrue($object->isSuperScript());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(true));
        $this->assertTrue($object->isSubScript());
        $this->assertFalse($object->isSuperScript());
    }

    /**
     * Test get/set superscript
     */
    public function testSetIsSuperScript()
    {
        $object = new Font();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript());
        $this->assertFalse($object->isSuperScript());
        $this->assertTrue($object->isSubScript());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(''));
        $this->assertFalse($object->isSuperScript());
        $this->assertTrue($object->isSubScript());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(false));
        $this->assertFalse($object->isSuperScript());
        $this->assertTrue($object->isSubScript());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(true));
        $this->assertTrue($object->isSuperScript());
        $this->assertFalse($object->isSubScript());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Font();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
}
