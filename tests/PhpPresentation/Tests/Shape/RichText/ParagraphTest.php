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

namespace PhpOffice\PhpPresentation\Tests\Shape\RichText;

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Font;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Paragraph element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\RichText\Paragraph
 */
class ParagraphTest extends TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $object = new Paragraph();
        $this->assertEmpty($object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->getAlignment());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->getBulletStyle());
    }

    public function testAlignment()
    {
        $object = new Paragraph();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->getAlignment());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setAlignment(new Alignment()));
    }

    /**
     * Test get/set bullet style
     */
    public function testBulletStyle()
    {
        $object = new Paragraph();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->getBulletStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setBulletStyle());
        $this->assertNull($object->getBulletStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setBulletStyle(new Bullet()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->getBulletStyle());
    }

    /**
     * Test get/set font
     */
    public function testFont()
    {
        $object = new Paragraph();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setFont(new Font()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
    }

    /**
     * Test get/set hashCode
     */
    public function testHashCode()
    {
        $object = new Paragraph();
        $oElement = new TextElement();
        $object->addText($oElement);
        $this->assertEquals(md5($oElement->getHashCode().$object->getFont()->getHashCode().get_class($object)), $object->getHashCode());
    }

    /**
     * Test get/set hashIndex
     */
    public function testHashIndex()
    {
        $object = new Paragraph();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set linespacing
     */
    public function testLineSpacing()
    {
        $object = new Paragraph();
        $valueExpected = mt_rand(1, 100);
        $this->assertEquals(100, $object->getLineSpacing());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setLineSpacing($valueExpected));
        $this->assertEquals($valueExpected, $object->getLineSpacing());
    }

    /**
     * Test get/set richTextElements
     */
    public function testRichTextElements()
    {
        $object = new Paragraph();
        $this->assertInternalType('array', $object->getRichTextElements());
        $this->assertEmpty($object->getRichTextElements());
        $object->createBreak();
        $this->assertCount(1, $object->getRichTextElements());

        $array = array(
            new TextElement(),
            new TextElement(),
            new TextElement(),
        );
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->setRichTextElements($array));
        $this->assertCount(3, $object->getRichTextElements());
    }

    /**
     * @expectedException \Exception
     * expectedExceptionMessage Invalid \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface[] array passed.
     */
    public function testRichTextElementsException()
    {
        $object = new Paragraph();
        $object->setRichTextElements(1);
    }

    /**
     * Test text methods
     */
    public function testText()
    {
        $object = new Paragraph();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $object->addText(new TextElement()));
        $this->assertCount(1, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->createText());
        $this->assertCount(2, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->createText('AAA'));
        $this->assertCount(3, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->createBreak());
        $this->assertCount(4, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun());
        $this->assertCount(5, $object->getRichTextElements());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->createTextRun('BBB'));
        $this->assertCount(6, $object->getRichTextElements());
        $this->assertEquals('AAA'."\r\n".'BBB', $object->getPlainText());
        $this->assertEquals('AAA'."\r\n".'BBB', (string) $object);
    }
}
