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
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class FontTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Font();
        static::assertEquals('Calibri', $object->getName());
        static::assertEquals(10, $object->getSize());
        static::assertFalse($object->isBold());
        static::assertFalse($object->isItalic());
        static::assertFalse($object->isSuperScript());
        static::assertFalse($object->isSubScript());
        static::assertFalse($object->isStrikethrough());
        static::assertEquals(Font::UNDERLINE_NONE, $object->getUnderline());
        static::assertEquals(0, $object->getCharacterSpacing());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        static::assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
    }

    /**
     * Test get/set color
     * @expectedException \Exception
     * @expectedExceptionMessage $pValue must be an instance of \PhpOffice\PhpPresentation\Style\Color
     */
    public function testSetGetColorException()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setColor());
    }

    /**
     * Test get/set Character Spacing
     */
    public function testSetGetCharacterSpacing()
    {
        $object = new Font();
        static::assertEquals(0, $object->getCharacterSpacing());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setCharacterSpacing(0));
        static::assertEquals(0, $object->getCharacterSpacing());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setCharacterSpacing(10));
        static::assertEquals(1000, $object->getCharacterSpacing());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setCharacterSpacing());
        static::assertEquals(0, $object->getCharacterSpacing());
    }

    /**
     * Test get/set color
     */
    public function testSetGetColor()
    {
        $object = new Font();
        static::assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setColor(new Color(Color::COLOR_BLUE)));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        static::assertEquals(Color::COLOR_BLUE, $object->getColor()->getARGB());
    }

    /**
     * Test get/set name
     */
    public function testSetGetName()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setName());
        static::assertEquals('Calibri', $object->getName());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setName(''));
        static::assertEquals('Calibri', $object->getName());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setName('Arial'));
        static::assertEquals('Arial', $object->getName());
    }

    /**
     * Test get/set size
     */
    public function testSetGetSize()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSize());
        static::assertEquals(10, $object->getSize());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSize(''));
        static::assertEquals(10, $object->getSize());
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSize($value));
        static::assertEquals($value, $object->getSize());
    }

    /**
     * Test get/set underline
     */
    public function testSetGetUnderline()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setUnderline());
        static::assertEquals(Font::UNDERLINE_NONE, $object->getUnderline());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setUnderline(''));
        static::assertEquals(Font::UNDERLINE_NONE, $object->getUnderline());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setUnderline(Font::UNDERLINE_DASH));
        static::assertEquals(Font::UNDERLINE_DASH, $object->getUnderline());
    }

    /**
     * Test get/set bold
     */
    public function testSetIsBold()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold());
        static::assertFalse($object->isBold());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold(''));
        static::assertFalse($object->isBold());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold(false));
        static::assertFalse($object->isBold());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setBold(true));
        static::assertTrue($object->isBold());
    }

    /**
     * Test get/set italic
     */
    public function testSetIsItalic()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic());
        static::assertFalse($object->isItalic());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic(''));
        static::assertFalse($object->isItalic());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic(false));
        static::assertFalse($object->isItalic());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setItalic(true));
        static::assertTrue($object->isItalic());
    }

    /**
     * Test get/set strikethrough
     */
    public function testSetIsStriketrough()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough());
        static::assertFalse($object->isStrikethrough());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough(''));
        static::assertFalse($object->isStrikethrough());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough(false));
        static::assertFalse($object->isStrikethrough());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setStrikethrough(true));
        static::assertTrue($object->isStrikethrough());
    }

    /**
     * Test get/set subscript
     */
    public function testSetIsSubScript()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript());
        static::assertFalse($object->isSubScript());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(''));
        static::assertFalse($object->isSubScript());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(false));
        static::assertFalse($object->isSubScript());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(true));
        static::assertTrue($object->isSubScript());

        // Test toggle of SubScript
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(false));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(false));
        static::assertFalse($object->isSubScript());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(true));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(true));
        static::assertFalse($object->isSubScript());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(true));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(false));
        static::assertTrue($object->isSubScript());
    }

    /**
     * Test get/set superscript
     */
    public function testSetIsSuperScript()
    {
        $object = new Font();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript());
        static::assertFalse($object->isSuperScript());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(''));
        static::assertFalse($object->isSuperScript());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(false));
        static::assertFalse($object->isSuperScript());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(true));
        static::assertTrue($object->isSuperScript());

        // Test toggle of SubScript
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(false));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(false));
        static::assertFalse($object->isSuperScript());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(true));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(true));
        static::assertFalse($object->isSuperScript());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSuperScript(true));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->setSubScript(false));
        static::assertTrue($object->isSuperScript());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Font();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        static::assertEquals($value, $object->getHashIndex());
    }
}
