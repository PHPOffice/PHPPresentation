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

namespace PhpOffice\PhpPresentation\Tests\Shape\Table;

use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\Style\Borders;
use PhpOffice\PhpPresentation\Style\Fill;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Cell element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Cell
 */
class CellTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testConstruct(): void
    {
        $object = new Cell();
        self::assertEquals(0, $object->getActiveParagraphIndex());
        self::assertCount(1, $object->getParagraphs());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Borders', $object->getBorders());
    }

    public function testActiveParagraph(): void
    {
        $object = new Cell();
        self::assertEquals(0, $object->getActiveParagraphIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->createParagraph());
        self::assertCount(2, $object->getParagraphs());
        $value = mt_rand(0, 1);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setActiveParagraph($value));
        self::assertEquals($value, $object->getActiveParagraphIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->getActiveParagraph());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->getParagraph());
        $value = mt_rand(0, 1);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->getParagraph($value));

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setParagraphs([]));
        self::assertCount(0, $object->getParagraphs());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->createParagraph());
        self::assertCount(1, $object->getParagraphs());
    }

    public function testActiveParagraphException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1000) is out of bounds (0, 1)');

        $object = new Cell();
        $object->setActiveParagraph(1000);
    }

    public function testGetParagraphException(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (1000) is out of bounds (0, 1)');

        $object = new Cell();
        $object->getParagraph(1000);
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Cell();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }

    public function testText(): void
    {
        $object = new Cell();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->addText());
        self::assertCount(1, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->addText(new TextElement()));
        self::assertCount(2, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->createText());
        self::assertCount(3, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->createText('ALPHA'));
        self::assertCount(4, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->createBreak());
        self::assertCount(5, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun());
        self::assertCount(6, $object->getActiveParagraph()->getRichTextElements());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun('BETA'));
        self::assertCount(7, $object->getActiveParagraph()->getRichTextElements());
        self::assertEquals('ALPHA' . "\r\n" . 'BETA', $object->getPlainText());
        self::assertEquals('ALPHA' . "\r\n" . 'BETA', (string) $object);
    }

    public function testParagraphs(): void
    {
        $object = new Cell();

        $array = [
            new Paragraph(),
            new Paragraph(),
            new Paragraph(),
        ];

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setParagraphs($array));
        self::assertCount(3, $object->getParagraphs());
        self::assertEquals(2, $object->getActiveParagraphIndex());
    }

    public function testGetSetBorders(): void
    {
        $object = new Cell();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setBorders(new Borders()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Borders', $object->getBorders());
    }

    public function testGetSetColspan(): void
    {
        $object = new Cell();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setColSpan());
        self::assertEquals(0, $object->getColSpan());

        $value = mt_rand(1, 100);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setColSpan($value));
        self::assertEquals($value, $object->getColSpan());
    }

    public function testGetSetFill(): void
    {
        $object = new Cell();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setFill(new Fill()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    public function testGetSetRowspan(): void
    {
        $object = new Cell();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setRowSpan());
        self::assertEquals(0, $object->getRowSpan());

        $value = mt_rand(1, 100);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setRowSpan($value));
        self::assertEquals($value, $object->getRowSpan());
    }

    public function testGetSetWidth(): void
    {
        $object = new Cell();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setWidth());
        self::assertEquals(0, $object->getWidth());

        $value = mt_rand(1, 100);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setWidth($value));
        self::assertEquals($value, $object->getWidth());
    }
}
