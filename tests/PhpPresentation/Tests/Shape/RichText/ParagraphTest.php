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

namespace PhpOffice\PhpPresentation\Tests\Shape\RichText;

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Font;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Paragraph element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
 */
class ParagraphTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testConstruct(): void
    {
        $object = new Paragraph();
        $this->assertEmpty($object->getRichTextElements());
        $this->assertInstanceOf(Alignment::class, $object->getAlignment());
        $this->assertInstanceOf(Font::class, $object->getFont());
        $this->assertInstanceOf(Bullet::class, $object->getBulletStyle());
    }

    public function testAlignment(): void
    {
        $object = new Paragraph();
        $this->assertInstanceOf(Alignment::class, $object->getAlignment());
        $this->assertInstanceOf(Paragraph::class, $object->setAlignment(new Alignment()));
    }

    /**
     * Test get/set bullet style.
     */
    public function testBulletStyle(): void
    {
        $object = new Paragraph();
        $this->assertInstanceOf(Bullet::class, $object->getBulletStyle());
        $this->assertInstanceOf(Paragraph::class, $object->setBulletStyle());
        $this->assertNull($object->getBulletStyle());
        $this->assertInstanceOf(Paragraph::class, $object->setBulletStyle(new Bullet()));
        $this->assertInstanceOf(Bullet::class, $object->getBulletStyle());
    }

    /**
     * Test get/set font.
     */
    public function testFont(): void
    {
        $object = new Paragraph();
        $this->assertInstanceOf(Font::class, $object->getFont());
        $this->assertInstanceOf(Paragraph::class, $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf(Paragraph::class, $object->setFont(new Font()));
        $this->assertInstanceOf(Font::class, $object->getFont());
    }

    /**
     * Test get/set hashCode.
     */
    public function testHashCode(): void
    {
        $object = new Paragraph();
        $oElement = new TextElement();
        $object->addText($oElement);
        $this->assertEquals(md5($oElement->getHashCode() . $object->getFont()->getHashCode() . get_class($object)), $object->getHashCode());
    }

    /**
     * Test get/set hashIndex.
     */
    public function testHashIndex(): void
    {
        $object = new Paragraph();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set linespacing.
     */
    public function testLineSpacing(): void
    {
        $object = new Paragraph();
        $valueExpected = mt_rand(1, 100);
        $this->assertEquals(100, $object->getLineSpacing());
        $this->assertInstanceOf(Paragraph::class, $object->setLineSpacing($valueExpected));
        $this->assertEquals($valueExpected, $object->getLineSpacing());
    }

    /**
     * Test text methods.
     */
    public function testLineSpacingMode(): void
    {
        $object = new Paragraph();
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $object->getLineSpacingMode());
        $this->assertInstanceOf(Paragraph::class, $object->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_POINT));
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_POINT, $object->getLineSpacingMode());
        $this->assertInstanceOf(Paragraph::class, $object->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_PERCENT));
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $object->getLineSpacingMode());
        $this->assertInstanceOf(Paragraph::class, $object->setLineSpacingMode('Unauthorized'));
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $object->getLineSpacingMode());
    }

    /**
     * Test get/set richTextElements.
     */
    public function testRichTextElements(): void
    {
        $object = new Paragraph();
        $this->assertIsArray($object->getRichTextElements());
        $this->assertEmpty($object->getRichTextElements());
        $object->createBreak();
        $this->assertCount(1, $object->getRichTextElements());

        $array = [
            new TextElement(),
            new TextElement(),
            new TextElement(),
        ];
        $this->assertInstanceOf(Paragraph::class, $object->setRichTextElements($array));
        $this->assertCount(3, $object->getRichTextElements());
    }

    public function testSpacingAfter(): void
    {
        $object = new Paragraph();
        $this->assertEquals(0, $object->getSpacingAfter());
        $this->assertInstanceOf(Paragraph::class, $object->setSpacingAfter(1));
        $this->assertEquals(1, $object->getSpacingAfter());
    }

    public function testSpacingBefore(): void
    {
        $object = new Paragraph();
        $this->assertEquals(0, $object->getSpacingBefore());
        $this->assertInstanceOf(Paragraph::class, $object->setSpacingBefore(1));
        $this->assertEquals(1, $object->getSpacingBefore());
    }

    /**
     * Test text methods.
     */
    public function testText(): void
    {
        $object = new Paragraph();
        $this->assertInstanceOf(Paragraph::class, $object->addText(new TextElement()));
        $this->assertCount(1, $object->getRichTextElements());
        $this->assertInstanceOf(TextElement::class, $object->createText());
        $this->assertCount(2, $object->getRichTextElements());
        $this->assertInstanceOf(TextElement::class, $object->createText('AAA'));
        $this->assertCount(3, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->createBreak());
        $this->assertCount(4, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun());
        $this->assertCount(5, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun('BBB'));
        $this->assertCount(6, $object->getRichTextElements());
        $this->assertEquals('AAA' . "\r\n" . 'BBB', $object->getPlainText());
        $this->assertEquals('AAA' . "\r\n" . 'BBB', (string) $object);
    }
}
