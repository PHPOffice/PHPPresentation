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
        $this->assertEquals(0, $object->getActiveParagraphIndex());
        $this->assertCount(1, $object->getParagraphs());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Borders', $object->getBorders());
    }

    public function testActiveParagraph(): void
    {
        $object = new Cell();
        $this->assertEquals(0, $object->getActiveParagraphIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->createParagraph());
        $this->assertCount(2, $object->getParagraphs());
        $value = mt_rand(0, 1);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setActiveParagraph($value));
        $this->assertEquals($value, $object->getActiveParagraphIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->getActiveParagraph());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->getParagraph());
        $value = mt_rand(0, 1);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->getParagraph($value));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setParagraphs([]));
        $this->assertCount(0, $object->getParagraphs());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->createParagraph());
        $this->assertCount(1, $object->getParagraphs());
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
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testText(): void
    {
        $object = new Cell();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->addText());
        $this->assertCount(1, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->addText(new TextElement()));
        $this->assertCount(2, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->createText());
        $this->assertCount(3, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->createText('ALPHA'));
        $this->assertCount(4, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->createBreak());
        $this->assertCount(5, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun());
        $this->assertCount(6, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun('BETA'));
        $this->assertCount(7, $object->getActiveParagraph()->getRichTextElements());
        $this->assertEquals('ALPHA' . "\r\n" . 'BETA', $object->getPlainText());
        $this->assertEquals('ALPHA' . "\r\n" . 'BETA', (string) $object);
    }

    public function testParagraphs(): void
    {
        $object = new Cell();

        $array = [
            new Paragraph(),
            new Paragraph(),
            new Paragraph(),
        ];

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setParagraphs($array));
        $this->assertCount(3, $object->getParagraphs());
        $this->assertEquals(2, $object->getActiveParagraphIndex());
    }

    public function testGetSetBorders(): void
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setBorders(new Borders()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Borders', $object->getBorders());
    }

    public function testGetSetColspan(): void
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setColSpan());
        $this->assertEquals(0, $object->getColSpan());

        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setColSpan($value));
        $this->assertEquals($value, $object->getColSpan());
    }

    public function testGetSetFill(): void
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setFill(new Fill()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    public function testGetSetRowspan(): void
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setRowSpan());
        $this->assertEquals(0, $object->getRowSpan());

        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setRowSpan($value));
        $this->assertEquals($value, $object->getRowSpan());
    }

    public function testGetSetWidth(): void
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());

        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }
}
