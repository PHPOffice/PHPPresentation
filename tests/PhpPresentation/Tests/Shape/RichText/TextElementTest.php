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

use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;

/**
 * Test class for TextElement element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\RichText\TextElement
 */
class TextElementTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $object = new TextElement();
        $this->assertEquals('', $object->getText());

        $object = new TextElement('AAA');
        $this->assertEquals('AAA', $object->getText());
    }

    public function testFont()
    {
        $object = new TextElement();
        $this->assertNull($object->getFont());
    }

    public function testHyperlink()
    {
        $object = new TextElement();
        $this->assertFalse($object->hasHyperlink());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setHyperlink());
        $this->assertFalse($object->hasHyperlink());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->getHyperlink());
        $this->assertTrue($object->hasHyperlink());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setHyperlink(new Hyperlink('http://www.google.fr')));
        $this->assertTrue($object->hasHyperlink());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->getHyperlink());
    }

    public function testLanguage()
    {
        $object = new TextElement();
        $this->assertNull($object->getLanguage());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setLanguage('en-US'));
        $this->assertEquals('en-US', $object->getLanguage());
    }

    public function testText()
    {
        $object = new TextElement();
        $this->assertEquals('', $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setText());
        $this->assertEquals('', $object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setText('AAA'));
        $this->assertEquals('AAA', $object->getText());
    }

    /**
     * Test get/set hash index
     */
    public function testHashCode()
    {
        $object = new TextElement();
        $this->assertEquals(md5(get_class($object)), $object->getHashCode());
    }
}
