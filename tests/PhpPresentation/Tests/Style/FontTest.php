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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\Exception\NotAllowedValueException;
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
        self::assertEquals('Calibri', $object->getName());
        self::assertEquals(10, $object->getSize());
        self::assertFalse($object->isBold());
        self::assertFalse($object->isItalic());
        self::assertFalse($object->isSuperScript());
        self::assertFalse($object->isSubScript());
        self::assertFalse($object->isStrikethrough());
        self::assertEquals(Font::UNDERLINE_NONE, $object->getUnderline());
        self::assertEquals(0, $object->getCharacterSpacing());
        self::assertEquals(Font::CAPITALIZATION_NONE, $object->getCapitalization());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        self::assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        self::assertEquals(0, $object->getBaseline());
    }

    /**
     * Test get/set Baseline.
     */
    public function testBaseline(): void
    {
        $object = new Font();
        self::assertEquals(0, $object->getBaseline());
        self::assertInstanceOf(Font::class, $object->setBaseline(Font::BASELINE_SUBSCRIPT));
        self::assertEquals(Font::BASELINE_SUBSCRIPT, $object->getBaseline());
    }

    /**
     * Test get/set Capitalization.
     */
    public function testCapitalization(): void
    {
        $object = new Font();
        self::assertEquals(Font::CAPITALIZATION_NONE, $object->getCapitalization());
        self::assertInstanceOf(Font::class, $object->setCapitalization(Font::CAPITALIZATION_ALL));
        self::assertEquals(Font::CAPITALIZATION_ALL, $object->getCapitalization());
        self::assertInstanceOf(Font::class, $object->setCapitalization());
        self::assertEquals(Font::CAPITALIZATION_NONE, $object->getCapitalization());
    }

    /**
     * Test get/set Capitalization exception.
     */
    public function testCapitalizationException(): void
    {
        $this->expectException(NotAllowedValueException::class);
        $this->expectExceptionMessage('The value "Invalid" is not in allowed values ("none", "all", "small")');

        $object = new Font();
        $object->setCapitalization('Invalid');
    }

    /**
     * Test get/set charset.
     */
    public function testCharset(): void
    {
        $object = new Font();
        self::assertEquals(Font::CHARSET_DEFAULT, $object->getCharset());
        self::assertInstanceOf(Font::class, $object->setCharset(12));
        self::assertEquals(12, $object->getCharset());
    }

    /**
     * Test get/set Character Spacing.
     */
    public function testCharacterSpacing(): void
    {
        $object = new Font();
        self::assertEquals(0, $object->getCharacterSpacing());
        self::assertInstanceOf(Font::class, $object->setCharacterSpacing(0));
        self::assertEquals(0, $object->getCharacterSpacing());
        self::assertInstanceOf(Font::class, $object->setCharacterSpacing(10));
        self::assertEquals(1000, $object->getCharacterSpacing());
        self::assertInstanceOf(Font::class, $object->setCharacterSpacing());
        self::assertEquals(0, $object->getCharacterSpacing());
    }

    /**
     * Test get/set color.
     */
    public function testColor(): void
    {
        $object = new Font();
        self::assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        self::assertInstanceOf(Font::class, $object->setColor(new Color(Color::COLOR_BLUE)));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        self::assertEquals(Color::COLOR_BLUE, $object->getColor()->getARGB());
    }

    /**
     * Test get/set name.
     */
    public function testFormat(): void
    {
        $object = new Font();
        self::assertEquals(Font::FORMAT_LATIN, $object->getFormat());
        self::assertInstanceOf(Font::class, $object->setFormat());
        self::assertEquals(Font::FORMAT_LATIN, $object->getFormat());
        self::assertInstanceOf(Font::class, $object->setFormat('UnAuthorized'));
        self::assertEquals(Font::FORMAT_LATIN, $object->getFormat());
        self::assertInstanceOf(Font::class, $object->setFormat(Font::FORMAT_EAST_ASIAN));
        self::assertEquals(Font::FORMAT_EAST_ASIAN, $object->getFormat());
        self::assertInstanceOf(Font::class, $object->setFormat(Font::FORMAT_COMPLEX_SCRIPT));
        self::assertEquals(Font::FORMAT_COMPLEX_SCRIPT, $object->getFormat());
    }

    /**
     * Test get/set name.
     */
    public function testName(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setName());
        self::assertEquals('Calibri', $object->getName());
        self::assertInstanceOf(Font::class, $object->setName(''));
        self::assertEquals('Calibri', $object->getName());
        self::assertInstanceOf(Font::class, $object->setName('Arial'));
        self::assertEquals('Arial', $object->getName());
    }

    /**
     * Test get/set panose.
     */
    public function testPanose(): void
    {
        $object = new Font();
        self::assertEquals('', $object->getPanose());
        self::assertInstanceOf(Font::class, $object->setPanose('4494D72242'));
        self::assertEquals('4494D72242', $object->getPanose());
        self::assertInstanceOf(Font::class, $object->setPanose('020B0604030504040204'));
        self::assertEquals('2B64354424', $object->getPanose());
    }

    /**
     * Test get/set panose (Exception : Invalid Length).
     */
    public function testPanoseExceptionInvalidLength(): void
    {
        $this->expectException(InvalidParameterException::class);
        $this->expectExceptionMessage('The parameter pValue can\'t have the value "12345" (Validation: The length is not correct)');

        $object = new Font();
        $object->setPanose('12345');
    }

    /**
     * Test get/set panose (Exception : Invalid Char).
     */
    public function testPanoseExceptionInvalidChar(): void
    {
        $this->expectException(InvalidParameterException::class);
        $this->expectExceptionMessage('The parameter pValue can\'t have the value "4494D7224X" (Validation: The character "X" is not allowed)');

        $object = new Font();
        $object->setPanose('4494D7224X');
    }

    /**
     * Test get/set pitch family.
     */
    public function testPitchFamily(): void
    {
        $object = new Font();
        self::assertEquals(0, $object->getPitchFamily());
        self::assertInstanceOf(Font::class, $object->setPitchFamily(12));
        self::assertEquals(12, $object->getPitchFamily());
    }

    /**
     * Test get/set size.
     */
    public function testSize(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setSize());
        self::assertEquals(10, $object->getSize());
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Font::class, $object->setSize($value));
        self::assertEquals($value, $object->getSize());
    }

    /**
     * Test get/set underline.
     */
    public function testUnderline(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setUnderline());
        self::assertEquals(Font::UNDERLINE_NONE, $object->getUnderline());
        self::assertInstanceOf(Font::class, $object->setUnderline(''));
        self::assertEquals(Font::UNDERLINE_NONE, $object->getUnderline());
        self::assertInstanceOf(Font::class, $object->setUnderline(Font::UNDERLINE_DASH));
        self::assertEquals(Font::UNDERLINE_DASH, $object->getUnderline());
    }

    /**
     * Test get/set bold.
     */
    public function testSetIsBold(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setBold());
        self::assertFalse($object->isBold());
        self::assertInstanceOf(Font::class, $object->setBold(false));
        self::assertFalse($object->isBold());
        self::assertInstanceOf(Font::class, $object->setBold(true));
        self::assertTrue($object->isBold());
    }

    /**
     * Test get/set italic.
     */
    public function testSetIsItalic(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setItalic());
        self::assertFalse($object->isItalic());
        self::assertInstanceOf(Font::class, $object->setItalic(false));
        self::assertFalse($object->isItalic());
        self::assertInstanceOf(Font::class, $object->setItalic(true));
        self::assertTrue($object->isItalic());
    }

    /**
     * Test get/set strikethrough.
     */
    public function testSetIsStriketrough(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setStrikethrough());
        self::assertFalse($object->isStrikethrough());
        self::assertEquals(Font::STRIKE_NONE, $object->getStrikethrough());
        // boolean
        self::assertInstanceOf(Font::class, $object->setStrikethrough(false));
        self::assertFalse($object->isStrikethrough());
        self::assertEquals(Font::STRIKE_NONE, $object->getStrikethrough());
        self::assertInstanceOf(Font::class, $object->setStrikethrough(true));
        self::assertTrue($object->isStrikethrough());
        self::assertEquals(Font::STRIKE_SINGLE, $object->getStrikethrough());
        // string
        self::assertInstanceOf(Font::class, $object->setStrikethrough(Font::STRIKE_DOUBLE));
        self::assertTrue($object->isStrikethrough());
        self::assertEquals(Font::STRIKE_DOUBLE, $object->getStrikethrough());
    }

    /**
     * Test get/set subscript.
     */
    public function testSetIsSubScript(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setSubScript());
        self::assertFalse($object->isSubScript());
        self::assertInstanceOf(Font::class, $object->setSubScript(false));
        self::assertFalse($object->isSubScript());
        self::assertInstanceOf(Font::class, $object->setSubScript(true));
        self::assertTrue($object->isSubScript());

        // Test toggle of SubScript
        self::assertInstanceOf(Font::class, $object->setSubScript(false));
        self::assertInstanceOf(Font::class, $object->setSuperScript(false));
        self::assertFalse($object->isSubScript());

        self::assertInstanceOf(Font::class, $object->setSubScript(true));
        self::assertInstanceOf(Font::class, $object->setSuperScript(true));
        self::assertFalse($object->isSubScript());

        self::assertInstanceOf(Font::class, $object->setSubScript(true));
        self::assertInstanceOf(Font::class, $object->setSuperScript(false));
        self::assertTrue($object->isSubScript());
    }

    /**
     * Test get/set superscript.
     */
    public function testSetIsSuperScript(): void
    {
        $object = new Font();
        self::assertInstanceOf(Font::class, $object->setSuperScript());
        self::assertFalse($object->isSuperScript());
        self::assertInstanceOf(Font::class, $object->setSuperScript(false));
        self::assertFalse($object->isSuperScript());
        self::assertInstanceOf(Font::class, $object->setSuperScript(true));
        self::assertTrue($object->isSuperScript());

        // Test toggle of SubScript
        self::assertInstanceOf(Font::class, $object->setSuperScript(false));
        self::assertInstanceOf(Font::class, $object->setSubScript(false));
        self::assertFalse($object->isSuperScript());

        self::assertInstanceOf(Font::class, $object->setSuperScript(true));
        self::assertInstanceOf(Font::class, $object->setSubScript(true));
        self::assertFalse($object->isSuperScript());

        self::assertInstanceOf(Font::class, $object->setSuperScript(true));
        self::assertInstanceOf(Font::class, $object->setSubScript(false));
        self::assertTrue($object->isSuperScript());
    }

    /**
     * Test get/set hash index.
     */
    public function testHashIndex(): void
    {
        $object = new Font();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }
}
