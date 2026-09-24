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
        self::assertEmpty($object->getRichTextElements());
        self::assertInstanceOf(Alignment::class, $object->getAlignment());
        self::assertInstanceOf(Font::class, $object->getFont());
        self::assertInstanceOf(Bullet::class, $object->getBulletStyle());
    }

    public function testAlignment(): void
    {
        $object = new Paragraph();
        self::assertInstanceOf(Alignment::class, $object->getAlignment());
        self::assertInstanceOf(Paragraph::class, $object->setAlignment(new Alignment()));
    }

    /**
     * Test get/set bullet style.
     */
    public function testBulletStyle(): void
    {
        $object = new Paragraph();
        self::assertInstanceOf(Bullet::class, $object->getBulletStyle());
        self::assertInstanceOf(Paragraph::class, $object->setBulletStyle());
        self::assertNull($object->getBulletStyle());
        self::assertInstanceOf(Paragraph::class, $object->setBulletStyle(new Bullet()));
        self::assertInstanceOf(Bullet::class, $object->getBulletStyle());
    }

    /**
     * Test get/set font.
     */
    public function testFont(): void
    {
        $object = new Paragraph();
        self::assertInstanceOf(Font::class, $object->getFont());
        // a font the setter was asked to take away is replaced with a fresh one, never a null
        self::assertInstanceOf(Paragraph::class, $object->setFont());
        self::assertInstanceOf(Font::class, $object->getFont());
        self::assertInstanceOf(Paragraph::class, $object->setFont(new Font()));
        self::assertInstanceOf(Font::class, $object->getFont());
    }

    /**
     * Test get/set hashCode.
     */
    public function testHashCode(): void
    {
        $object = new Paragraph();
        $oElement = new TextElement();
        $object->addText($oElement);
        self::assertEquals(md5($oElement->getHashCode() . $object->getFont()->getHashCode() . get_class($object)), $object->getHashCode());
    }

    /**
     * Test get/set hashIndex.
     */
    public function testHashIndex(): void
    {
        $object = new Paragraph();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set linespacing.
     */
    public function testLineSpacing(): void
    {
        $object = new Paragraph();
        $valueExpected = mt_rand(1, 100);
        self::assertEquals(100, $object->getLineSpacing());
        self::assertInstanceOf(Paragraph::class, $object->setLineSpacing($valueExpected));
        self::assertEquals($valueExpected, $object->getLineSpacing());
    }

    /**
     * Test text methods.
     */
    public function testLineSpacingMode(): void
    {
        $object = new Paragraph();
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $object->getLineSpacingMode());
        self::assertInstanceOf(Paragraph::class, $object->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_POINT));
        self::assertEquals(Paragraph::LINE_SPACING_MODE_POINT, $object->getLineSpacingMode());
        self::assertInstanceOf(Paragraph::class, $object->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_PERCENT));
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $object->getLineSpacingMode());
        self::assertInstanceOf(Paragraph::class, $object->setLineSpacingMode('Unauthorized'));
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $object->getLineSpacingMode());
    }

    /**
     * Test get/set richTextElements.
     */
    public function testRichTextElements(): void
    {
        $object = new Paragraph();
        self::assertEmpty($object->getRichTextElements());
        $object->createBreak();
        self::assertCount(1, $object->getRichTextElements());

        $array = [
            new TextElement(),
            new TextElement(),
            new TextElement(),
        ];
        self::assertInstanceOf(Paragraph::class, $object->setRichTextElements($array));
        self::assertCount(3, $object->getRichTextElements());
    }

    public function testSpacingAfter(): void
    {
        $object = new Paragraph();
        self::assertEquals(0, $object->getSpacingAfter());
        self::assertInstanceOf(Paragraph::class, $object->setSpacingAfter(1));
        self::assertEquals(1, $object->getSpacingAfter());
    }

    public function testSpacingBefore(): void
    {
        $object = new Paragraph();
        self::assertEquals(0, $object->getSpacingBefore());
        self::assertInstanceOf(Paragraph::class, $object->setSpacingBefore(1));
        self::assertEquals(1, $object->getSpacingBefore());
    }

    /**
     * Test text methods.
     */
    public function testText(): void
    {
        $object = new Paragraph();
        self::assertInstanceOf(Paragraph::class, $object->addText(new TextElement()));
        self::assertCount(1, $object->getRichTextElements());
        self::assertInstanceOf(TextElement::class, $object->createText());
        self::assertCount(2, $object->getRichTextElements());
        self::assertInstanceOf(TextElement::class, $object->createText('AAA'));
        self::assertCount(3, $object->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->createBreak());
        self::assertCount(4, $object->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun());
        self::assertCount(5, $object->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun('BBB'));
        self::assertCount(6, $object->getRichTextElements());
        self::assertEquals('AAA' . "\r\n" . 'BBB', $object->getPlainText());
        self::assertEquals('AAA' . "\r\n" . 'BBB', (string) $object);
    }

    public function testNumberingIsContinuedOverDeeperParagraphsAndEndedByShallowerOnesAndGaps(): void
    {
        // level, marker, start: a numbering from 2 with a deeper one inside it, a new one
        // from 3, a paragraph with no marker, and one more after it
        $paragraphs = [];
        foreach ([[0, 2], [0, null], [1, null], [1, null], [0, 3], [0, null], [0, false], [0, null]] as [$level, $startAt]) {
            $paragraph = new Paragraph();
            $paragraph->getAlignment()->setLevel($level);
            $paragraph->getBulletStyle()
                ->setBulletType(false === $startAt ? Bullet::TYPE_NONE : Bullet::TYPE_NUMERIC)
                ->setBulletNumericStartAt(false === $startAt ? null : $startAt);
            $paragraphs[] = $paragraph;
        }

        self::assertSame(
            [0 => null, 1 => 2, 2 => null, 3 => 1, 4 => 2, 5 => 3, 7 => null],
            Paragraph::getContinuedNumericStartAts($paragraphs)
        );
    }

    public function testNumberingIsEndedByABulletOrAnotherSchemeAtItsLevel(): void
    {
        $paragraphs = [];
        foreach ([[Bullet::TYPE_NUMERIC, Bullet::NUMERIC_ARABICPERIOD], [Bullet::TYPE_NUMERIC, Bullet::NUMERIC_ALPHALCPERIOD], [Bullet::TYPE_BULLET, Bullet::NUMERIC_ALPHALCPERIOD], [Bullet::TYPE_NUMERIC, Bullet::NUMERIC_ALPHALCPERIOD]] as [$type, $scheme]) {
            $paragraph = new Paragraph();
            $paragraph->getBulletStyle()->setBulletType($type)->setBulletNumericStyle($scheme);
            $paragraphs[] = $paragraph;
        }
        $paragraphs[0]->getBulletStyle()->setBulletNumericStartAt(2);

        self::assertSame([0 => null, 1 => null, 3 => null], Paragraph::getContinuedNumericStartAts($paragraphs));
    }

    public function testNumberingAskedToContinueTakesTheNumberAfterTheLastOne(): void
    {
        // 1, 2, a paragraph with no marker, 3, 4 asked to continue, another, then 1
        $paragraphs = [];
        foreach ([[true, false], [true, false], [false, false], [true, true], [true, false], [false, false], [true, false]] as [$numbered, $continue]) {
            $paragraph = new Paragraph();
            $paragraph->getBulletStyle()
                ->setBulletType($numbered ? Bullet::TYPE_NUMERIC : Bullet::TYPE_NONE)
                ->setBulletNumericContinue($continue);
            $paragraphs[] = $paragraph;
        }

        self::assertSame([0 => null, 1 => 1, 3 => 3, 4 => 3, 6 => null], Paragraph::getContinuedNumericStartAts($paragraphs));
        self::assertSame([0 => null, 3 => 3, 6 => 5], Paragraph::getResumedNumericStartAts($paragraphs));
    }
}
