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
        $this->assertEquals(0, $object->getActiveParagraphIndex());
        $this->assertCount(1, $object->getParagraphs());
    }

    public function testActiveParagraph(): void
    {
        $object = new RichText();
        $this->assertEquals(0, $object->getActiveParagraphIndex());
        $this->assertInstanceOf(Paragraph::class, $object->createParagraph());
        $this->assertCount(2, $object->getParagraphs());
        $value = mt_rand(0, 1);
        $this->assertInstanceOf(Paragraph::class, $object->setActiveParagraph($value));
        $this->assertEquals($value, $object->getActiveParagraphIndex());
        $this->assertInstanceOf(Paragraph::class, $object->getActiveParagraph());
        $this->assertInstanceOf(Paragraph::class, $object->getParagraph());
        $value = mt_rand(0, 1);
        $this->assertInstanceOf(Paragraph::class, $object->getParagraph($value));
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
        $this->assertInstanceOf(RichText::class, $object->setColumns($value));
        $this->assertEquals($value, $object->getColumns());
    }

    public function testColumnSpacing(): void
    {
        $object = new RichText();

        $this->assertEquals(0, $object->getColumnSpacing());
        $value = mt_rand(1, 16);
        $this->assertInstanceOf(RichText::class, $object->setColumnSpacing($value));
        $this->assertEquals($value, $object->getColumnSpacing());
        $this->assertInstanceOf(RichText::class, $object->setColumnSpacing(-1));
        $this->assertEquals($value, $object->getColumnSpacing());
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

        $this->assertInstanceOf(RichText::class, $object->setParagraphs($array));
        $this->assertCount(3, $object->getParagraphs());
        $this->assertEquals(2, $object->getActiveParagraphIndex());
    }

    public function testText(): void
    {
        $object = new RichText();
        $this->assertInstanceOf(RichText::class, $object->addText());
        $this->assertCount(1, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf(RichText::class, $object->addText(new TextElement()));
        $this->assertCount(2, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf(TextElement::class, $object->createText());
        $this->assertCount(3, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf(TextElement::class, $object->createText('ALPHA'));
        $this->assertCount(4, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf(BreakElement::class, $object->createBreak());
        $this->assertCount(5, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf(Run::class, $object->createTextRun());
        $this->assertCount(6, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf(Run::class, $object->createTextRun('BETA'));
        $this->assertCount(7, $object->getActiveParagraph()->getRichTextElements());
        $this->assertEquals('ALPHA' . "\r\n" . 'BETA', $object->getPlainText());
        $this->assertEquals('ALPHA' . "\r\n" . 'BETA', (string) $object);
    }

    public function testGetSetAutoFit(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setAutoFit());
        $this->assertEquals(RichText::AUTOFIT_DEFAULT, $object->getAutoFit());

        $this->assertInstanceOf(RichText::class, $object->setAutoFit(RichText::AUTOFIT_NORMAL));
        $this->assertEquals(RichText::AUTOFIT_NORMAL, $object->getAutoFit());
    }

    public function testGetSetHAutoShrink(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal());
        $this->assertNull($object->hasAutoShrinkHorizontal());

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal(null));
        $this->assertNull($object->hasAutoShrinkHorizontal());

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal(true));
        $this->assertTrue($object->hasAutoShrinkHorizontal());

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkHorizontal(false));
        $this->assertFalse($object->hasAutoShrinkHorizontal());
    }

    public function testGetSetVAutoShrink(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkVertical());
        $this->assertNull($object->hasAutoShrinkVertical());

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkVertical(null));
        $this->assertNull($object->hasAutoShrinkVertical());

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkVertical(true));
        $this->assertTrue($object->hasAutoShrinkVertical());

        $this->assertInstanceOf(RichText::class, $object->setAutoShrinkVertical(false));
        $this->assertFalse($object->hasAutoShrinkVertical());
    }

    public function testGetSetHOverflow(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setHorizontalOverflow());
        $this->assertEquals(RichText::OVERFLOW_OVERFLOW, $object->getHorizontalOverflow());

        $this->assertInstanceOf(RichText::class, $object->setHorizontalOverflow(RichText::OVERFLOW_CLIP));
        $this->assertEquals(RichText::OVERFLOW_CLIP, $object->getHorizontalOverflow());
    }

    public function testGetSetInset(): void
    {
        $object = new RichText();

        // Default
        $this->assertInstanceOf(RichText::class, $object->setInsetBottom());
        $this->assertEquals(4.8, $object->getInsetBottom());
        $this->assertInstanceOf(RichText::class, $object->setInsetLeft());
        $this->assertEquals(9.6, $object->getInsetLeft());
        $this->assertInstanceOf(RichText::class, $object->setInsetRight());
        $this->assertEquals(9.6, $object->getInsetRight());
        $this->assertInstanceOf(RichText::class, $object->setInsetTop());
        $this->assertEquals(4.8, $object->getInsetTop());

        // Value
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(RichText::class, $object->setInsetBottom($value));
        $this->assertEquals($value, $object->getInsetBottom());
        $this->assertInstanceOf(RichText::class, $object->setInsetLeft($value));
        $this->assertEquals($value, $object->getInsetLeft());
        $this->assertInstanceOf(RichText::class, $object->setInsetRight($value));
        $this->assertEquals($value, $object->getInsetRight());
        $this->assertInstanceOf(RichText::class, $object->setInsetTop($value));
        $this->assertEquals($value, $object->getInsetTop());
    }

    public function testGetSetUpright(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setUpright());
        $this->assertFalse($object->isUpright());

        $this->assertInstanceOf(RichText::class, $object->setUpright(true));
        $this->assertTrue($object->isUpright());

        $this->assertInstanceOf(RichText::class, $object->setUpright(false));
        $this->assertFalse($object->isUpright());
    }

    public function testGetSetVertical(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setVertical());
        $this->assertFalse($object->isVertical());

        $this->assertInstanceOf(RichText::class, $object->setVertical(true));
        $this->assertTrue($object->isVertical());

        $this->assertInstanceOf(RichText::class, $object->setVertical(false));
        $this->assertFalse($object->isVertical());
    }

    public function testGetSetVOverflow(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setVerticalOverflow());
        $this->assertEquals(RichText::OVERFLOW_OVERFLOW, $object->getVerticalOverflow());

        $this->assertInstanceOf(RichText::class, $object->setVerticalOverflow(RichText::OVERFLOW_CLIP));
        $this->assertEquals(RichText::OVERFLOW_CLIP, $object->getVerticalOverflow());
    }

    public function testGetSetWrap(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(RichText::class, $object->setWrap());
        $this->assertEquals(RichText::WRAP_SQUARE, $object->getWrap());

        $this->assertInstanceOf(RichText::class, $object->setWrap(RichText::WRAP_NONE));
        $this->assertEquals(RichText::WRAP_NONE, $object->getWrap());
    }

    public function testHashCode(): void
    {
        $object = new RichText();

        $hash = $object->getActiveParagraph()->getHashCode();
        $hash .= RichText::WRAP_SQUARE . RichText::AUTOFIT_DEFAULT . RichText::OVERFLOW_OVERFLOW . RichText::OVERFLOW_OVERFLOW . '00104.89.69.64.8';
        $hash .= md5('00000' . $object->getFill()->getHashCode() . $object->getShadow()->getHashCode() . '' . get_parent_class($object));
        $hash .= get_class($object);
        $this->assertEquals(md5($hash), $object->getHashCode());
    }
}
