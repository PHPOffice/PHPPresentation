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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class FontTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
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
     * Test get/set Character Spacing.
     */
    public function testCharacterSpacing(): void
    {
        $object = new Font();
        $this->assertEquals(0, $object->getCharacterSpacing());
        $this->assertInstanceOf(Font::class, $object->setCharacterSpacing(0));
        $this->assertEquals(0, $object->getCharacterSpacing());
        $this->assertInstanceOf(Font::class, $object->setCharacterSpacing(10));
        $this->assertEquals(1000, $object->getCharacterSpacing());
        $this->assertInstanceOf(Font::class, $object->setCharacterSpacing());
        $this->assertEquals(0, $object->getCharacterSpacing());
    }

    /**
     * Test get/set color.
     */
    public function testColor(): void
    {
        $object = new Font();
        $this->assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        $this->assertInstanceOf(Font::class, $object->setColor(new Color(Color::COLOR_BLUE)));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertEquals(Color::COLOR_BLUE, $object->getColor()->getARGB());
    }

    /**
     * Test get/set name.
     */
    public function testFormat(): void
    {
        $object = new Font();
        $this->assertEquals(Font::FORMAT_LATIN, $object->getFormat());
        $this->assertInstanceOf(Font::class, $object->setFormat());
        $this->assertEquals(Font::FORMAT_LATIN, $object->getFormat());
        $this->assertInstanceOf(Font::class, $object->setFormat('UnAuthorized'));
        $this->assertEquals(Font::FORMAT_LATIN, $object->getFormat());
        $this->assertInstanceOf(Font::class, $object->setFormat(Font::FORMAT_EAST_ASIAN));
        $this->assertEquals(Font::FORMAT_EAST_ASIAN, $object->getFormat());
        $this->assertInstanceOf(Font::class, $object->setFormat(Font::FORMAT_COMPLEX_SCRIPT));
        $this->assertEquals(Font::FORMAT_COMPLEX_SCRIPT, $object->getFormat());
    }

    /**
     * Test get/set name.
     */
    public function testName(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setName());
        $this->assertEquals('Calibri', $object->getName());
        $this->assertInstanceOf(Font::class, $object->setName(''));
        $this->assertEquals('Calibri', $object->getName());
        $this->assertInstanceOf(Font::class, $object->setName('Arial'));
        $this->assertEquals('Arial', $object->getName());
    }

    /**
     * Test get/set size.
     */
    public function testSize(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setSize());
        $this->assertEquals(10, $object->getSize());
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Font::class, $object->setSize($value));
        $this->assertEquals($value, $object->getSize());
    }

    /**
     * Test get/set underline.
     */
    public function testUnderline(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setUnderline());
        $this->assertEquals(FONT::UNDERLINE_NONE, $object->getUnderline());
        $this->assertInstanceOf(Font::class, $object->setUnderline(''));
        $this->assertEquals(FONT::UNDERLINE_NONE, $object->getUnderline());
        $this->assertInstanceOf(Font::class, $object->setUnderline(FONT::UNDERLINE_DASH));
        $this->assertEquals(FONT::UNDERLINE_DASH, $object->getUnderline());
    }

    /**
     * Test get/set bold.
     */
    public function testSetIsBold(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setBold());
        $this->assertFalse($object->isBold());
        $this->assertInstanceOf(Font::class, $object->setBold(false));
        $this->assertFalse($object->isBold());
        $this->assertInstanceOf(Font::class, $object->setBold(true));
        $this->assertTrue($object->isBold());
    }

    /**
     * Test get/set italic.
     */
    public function testSetIsItalic(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setItalic());
        $this->assertFalse($object->isItalic());
        $this->assertInstanceOf(Font::class, $object->setItalic(false));
        $this->assertFalse($object->isItalic());
        $this->assertInstanceOf(Font::class, $object->setItalic(true));
        $this->assertTrue($object->isItalic());
    }

    /**
     * Test get/set strikethrough.
     */
    public function testSetIsStriketrough(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setStrikethrough());
        $this->assertFalse($object->isStrikethrough());
        $this->assertInstanceOf(Font::class, $object->setStrikethrough(false));
        $this->assertFalse($object->isStrikethrough());
        $this->assertInstanceOf(Font::class, $object->setStrikethrough(true));
        $this->assertTrue($object->isStrikethrough());
    }

    /**
     * Test get/set subscript.
     */
    public function testSetIsSubScript(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setSubScript());
        $this->assertFalse($object->isSubScript());
        $this->assertInstanceOf(Font::class, $object->setSubScript(false));
        $this->assertFalse($object->isSubScript());
        $this->assertInstanceOf(Font::class, $object->setSubScript(true));
        $this->assertTrue($object->isSubScript());

        // Test toggle of SubScript
        $this->assertInstanceOf(Font::class, $object->setSubScript(false));
        $this->assertInstanceOf(Font::class, $object->setSuperScript(false));
        $this->assertFalse($object->isSubScript());

        $this->assertInstanceOf(Font::class, $object->setSubScript(true));
        $this->assertInstanceOf(Font::class, $object->setSuperScript(true));
        $this->assertFalse($object->isSubScript());

        $this->assertInstanceOf(Font::class, $object->setSubScript(true));
        $this->assertInstanceOf(Font::class, $object->setSuperScript(false));
        $this->assertTrue($object->isSubScript());
    }

    /**
     * Test get/set superscript.
     */
    public function testSetIsSuperScript(): void
    {
        $object = new Font();
        $this->assertInstanceOf(Font::class, $object->setSuperScript());
        $this->assertFalse($object->isSuperScript());
        $this->assertInstanceOf(Font::class, $object->setSuperScript(false));
        $this->assertFalse($object->isSuperScript());
        $this->assertInstanceOf(Font::class, $object->setSuperScript(true));
        $this->assertTrue($object->isSuperScript());

        // Test toggle of SubScript
        $this->assertInstanceOf(Font::class, $object->setSuperScript(false));
        $this->assertInstanceOf(Font::class, $object->setSubScript(false));
        $this->assertFalse($object->isSuperScript());

        $this->assertInstanceOf(Font::class, $object->setSuperScript(true));
        $this->assertInstanceOf(Font::class, $object->setSubScript(true));
        $this->assertFalse($object->isSuperScript());

        $this->assertInstanceOf(Font::class, $object->setSuperScript(true));
        $this->assertInstanceOf(Font::class, $object->setSubScript(false));
        $this->assertTrue($object->isSuperScript());
    }

    /**
     * Test get/set hash index.
     */
    public function testHashIndex(): void
    {
        $object = new Font();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
}
