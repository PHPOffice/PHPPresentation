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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PHPUnit\Framework\TestCase;

/**
 * Test class for RichText element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\RichText
 */
class RichTextTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new RichText();
        self::assertEquals(0, $object->getActiveParagraphIndex());
        self::assertCount(1, $object->getParagraphs());
    }

    public function testActiveParagraph(): void
    {
        $object = new RichText();
        self::assertEquals(0, $object->getActiveParagraphIndex());
        self::assertInstanceOf(Paragraph::class, $object->createParagraph());
        self::assertCount(2, $object->getParagraphs());
        $value = mt_rand(0, 1);
        self::assertInstanceOf(Paragraph::class, $object->setActiveParagraph($value));
        self::assertEquals($value, $object->getActiveParagraphIndex());
        self::assertInstanceOf(Paragraph::class, $object->getActiveParagraph());
        self::assertInstanceOf(Paragraph::class, $object->getParagraph());
        $value = mt_rand(0, 1);
        self::assertInstanceOf(Paragraph::class, $object->getParagraph($value));
    }

    public function testActiveParagraphException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1000) is out of bounds (0, 1)');

        $object = new RichText();
        $object->setActiveParagraph(1000);
    }

    public function testGetParagraphException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1000) is out of bounds (0, 1)');

        $object = new RichText();
        $object->getParagraph(1000);
    }

    public function testColumns(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 16);
        self::assertInstanceOf(RichText::class, $object->setColumns($value));
        self::assertEquals($value, $object->getColumns());
    }

    public function testColumnSpacing(): void
    {
        $object = new RichText();

        self::assertEquals(0, $object->getColumnSpacing());
        $value = mt_rand(1, 16);
        self::assertInstanceOf(RichText::class, $object->setColumnSpacing($value));
        self::assertEquals($value, $object->getColumnSpacing());
        self::assertInstanceOf(RichText::class, $object->setColumnSpacing(-1));
        self::assertEquals($value, $object->getColumnSpacing());
    }

    public function testColumnsException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1000) is out of bounds (1, 16)');

        $object = new RichText();
        $object->setColumns(1000);
    }

    public function testParagraphs(): void
    {
        $object = new RichText();

        $array = [
            new Paragraph(),
            new Paragraph(),
            new Paragraph(),
        ];

        self::assertInstanceOf(RichText::class, $object->setParagraphs($array));
        self::assertCount(3, $object->getParagraphs());
        self::assertEquals(2, $object->getActiveParagraphIndex());
    }

    public function testText(): void
    {
        $object = new RichText();
        self::assertInstanceOf(RichText::class, $object->addText());
        self::assertCount(1, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf(RichText::class, $object->addText(new TextElement()));
        self::assertCount(2, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf(TextElement::class, $object->createText());
        self::assertCount(3, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf(TextElement::class, $object->createText('ALPHA'));
        self::assertCount(4, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf(BreakElement::class, $object->createBreak());
        self::assertCount(5, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf(Run::class, $object->createTextRun());
        self::assertCount(6, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf(Run::class, $object->createTextRun('BETA'));
        self::assertCount(7, $object->getActiveParagraph()->getRichTextElements());
        self::assertEquals('ALPHA' . "\r\n" . 'BETA', $object->getPlainText());
        self::assertEquals('ALPHA' . "\r\n" . 'BETA', (string) $object);
    }

    public function testGetSetAutoFit(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setAutoFit());
        self::assertEquals(RichText::AUTOFIT_DEFAULT, $object->getAutoFit());

        self::assertInstanceOf(RichText::class, $object->setAutoFit(RichText::AUTOFIT_NORMAL));
        self::assertEquals(RichText::AUTOFIT_NORMAL, $object->getAutoFit());
    }

    public function testGetSetHAutoShrink(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal());
        self::assertNull($object->hasAutoShrinkHorizontal());

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal(null));
        self::assertNull($object->hasAutoShrinkHorizontal());

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal(true));
        self::assertTrue($object->hasAutoShrinkHorizontal());

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal(false));
        self::assertFalse($object->hasAutoShrinkHorizontal());
    }

    public function testGetSetVAutoShrink(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkVertical());
        self::assertNull($object->hasAutoShrinkVertical());

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkVertical(null));
        self::assertNull($object->hasAutoShrinkVertical());

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkVertical(true));
        self::assertTrue($object->hasAutoShrinkVertical());

        self::assertInstanceOf(RichText::class, $object->setAutoShrinkVertical(false));
        self::assertFalse($object->hasAutoShrinkVertical());
    }

    public function testGetSetHOverflow(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setHorizontalOverflow());
        self::assertEquals(RichText::OVERFLOW_OVERFLOW, $object->getHorizontalOverflow());

        self::assertInstanceOf(RichText::class, $object->setHorizontalOverflow(RichText::OVERFLOW_CLIP));
        self::assertEquals(RichText::OVERFLOW_CLIP, $object->getHorizontalOverflow());
    }

    public function testGetSetInset(): void
    {
        $object = new RichText();

        // Default
        self::assertInstanceOf(RichText::class, $object->setInsetBottom());
        self::assertEquals(4.8, $object->getInsetBottom());
        self::assertInstanceOf(RichText::class, $object->setInsetLeft());
        self::assertEquals(9.6, $object->getInsetLeft());
        self::assertInstanceOf(RichText::class, $object->setInsetRight());
        self::assertEquals(9.6, $object->getInsetRight());
        self::assertInstanceOf(RichText::class, $object->setInsetTop());
        self::assertEquals(4.8, $object->getInsetTop());

        // Value
        $value = mt_rand(1, 100);
        self::assertInstanceOf(RichText::class, $object->setInsetBottom($value));
        self::assertEquals($value, $object->getInsetBottom());
        self::assertInstanceOf(RichText::class, $object->setInsetLeft($value));
        self::assertEquals($value, $object->getInsetLeft());
        self::assertInstanceOf(RichText::class, $object->setInsetRight($value));
        self::assertEquals($value, $object->getInsetRight());
        self::assertInstanceOf(RichText::class, $object->setInsetTop($value));
        self::assertEquals($value, $object->getInsetTop());
    }

    public function testGetSetUpright(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setUpright());
        self::assertFalse($object->isUpright());

        self::assertInstanceOf(RichText::class, $object->setUpright(true));
        self::assertTrue($object->isUpright());

        self::assertInstanceOf(RichText::class, $object->setUpright(false));
        self::assertFalse($object->isUpright());
    }

    public function testGetSetVertical(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setVertical());
        self::assertFalse($object->isVertical());

        self::assertInstanceOf(RichText::class, $object->setVertical(true));
        self::assertTrue($object->isVertical());

        self::assertInstanceOf(RichText::class, $object->setVertical(false));
        self::assertFalse($object->isVertical());
    }

    public function testGetSetVOverflow(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setVerticalOverflow());
        self::assertEquals(RichText::OVERFLOW_OVERFLOW, $object->getVerticalOverflow());

        self::assertInstanceOf(RichText::class, $object->setVerticalOverflow(RichText::OVERFLOW_CLIP));
        self::assertEquals(RichText::OVERFLOW_CLIP, $object->getVerticalOverflow());
    }

    public function testGetSetWrap(): void
    {
        $object = new RichText();

        self::assertInstanceOf(RichText::class, $object->setWrap());
        self::assertEquals(RichText::WRAP_SQUARE, $object->getWrap());

        self::assertInstanceOf(RichText::class, $object->setWrap(RichText::WRAP_NONE));
        self::assertEquals(RichText::WRAP_NONE, $object->getWrap());
    }

    public function testHashCode(): void
    {
        $object = new RichText();

        $hash = $object->getActiveParagraph()->getHashCode();
        $hash .= RichText::WRAP_SQUARE . RichText::AUTOFIT_DEFAULT . RichText::OVERFLOW_OVERFLOW . RichText::OVERFLOW_OVERFLOW . '00104.89.69.64.80';
        $hash .= md5('00000' . $object->getFill()->getHashCode() . $object->getShadow()->getHashCode() . '' . get_parent_class($object));
        $hash .= get_class($object);
        self::assertEquals(md5($hash), $object->getHashCode());
    }
}
