<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shape\Table;

use PhpOffice\PhpPowerpoint\Shape\Table\Cell;
use PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph;
use PhpOffice\PhpPowerpoint\Shape\RichText\TextElement;
use PhpOffice\PhpPowerpoint\Style\Borders;
use PhpOffice\PhpPowerpoint\Style\Fill;

/**
 * Test class for Cell element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Cell
 */
class CellTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $object = new Cell();
        $this->assertEquals(0, $object->getActiveParagraphIndex());
        $this->assertCount(1, $object->getParagraphs());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Fill', $object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Borders', $object->getBorders());
    }

    public function testActiveParagraph()
    {
        $object = new Cell();
        $this->assertEquals(0, $object->getActiveParagraphIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Paragraph', $object->createParagraph());
        $this->assertCount(2, $object->getParagraphs());
        $value = rand(0, 1);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Paragraph', $object->setActiveParagraph($value));
        $this->assertEquals($value, $object->getActiveParagraphIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Paragraph', $object->getActiveParagraph());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Paragraph', $object->getParagraph());
        $value = rand(0, 1);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Paragraph', $object->getParagraph($value));
    }

    /**
     * @expectedException \Exception
     * expectedExceptionMessage Invalid paragraph count.
     */
    public function testActiveParagraphException()
    {
        $object = new Cell();
        $object->setActiveParagraph(1000);
    }

    /**
     * @expectedException \Exception
     * expectedExceptionMessage Invalid paragraph count.
     */
    public function testGetParagraphException()
    {
        $object = new Cell();
        $object->getParagraph(1000);
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Cell();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testText()
    {
        $object = new Cell();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->addText());
        $this->assertCount(1, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->addText(new TextElement()));
        $this->assertCount(2, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\TextElement', $object->createText());
        $this->assertCount(3, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\TextElement', $object->createText('ALPHA'));
        $this->assertCount(4, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\BreakElement', $object->createBreak());
        $this->assertCount(5, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Run', $object->createTextRun());
        $this->assertCount(6, $object->getActiveParagraph()->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\RichText\\Run', $object->createTextRun('BETA'));
        $this->assertCount(7, $object->getActiveParagraph()->getRichTextElements());
        $this->assertEquals('ALPHA'."\r\n".'BETA', $object->getPlainText());
        $this->assertEquals('ALPHA'."\r\n".'BETA', (string) $object);
    }

    public function testParagraphs()
    {
        $object = new Cell();

        $array = array(
            new Paragraph(),
            new Paragraph(),
            new Paragraph(),
        );

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setParagraphs($array));
        $this->assertCount(3, $object->getParagraphs());
        $this->assertEquals(2, $object->getActiveParagraphIndex());
    }

    /**
     * @expectedException \Exception
     * expectedExceptionMessage Invalid \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph[] array passed.
     */
    public function testParagraphsException()
    {
        $object = new Cell();
        $object->setParagraphs(1000);
    }

    public function testGetSetBorders()
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setBorders(new Borders()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Borders', $object->getBorders());
    }

    public function testGetSetColspan()
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setColSpan());
        $this->assertEquals(0, $object->getColSpan());

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setColSpan($value));
        $this->assertEquals($value, $object->getColSpan());
    }

    public function testGetSetFill()
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setFill(new Fill()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Fill', $object->getFill());
    }

    public function testGetSetRowspan()
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setRowSpan());
        $this->assertEquals(0, $object->getRowSpan());

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setRowSpan($value));
        $this->assertEquals($value, $object->getRowSpan());
    }

    public function testGetSetWidth()
    {
        $object = new Cell();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Table\\Cell', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }
}
