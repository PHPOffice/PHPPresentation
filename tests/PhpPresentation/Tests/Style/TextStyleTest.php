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

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\TextStyle;

class TextStyleTest extends \PHPUnit_Framework_TestCase
{
    public function testConstructDefaultTrue()
    {
        /**
         * @var \PhpOffice\PhpPresentation\Shape\RichText\Paragraph $oParagraph
         */
        $object = new TextStyle();

        $arrayBodyStyle = $object->getBodyStyle();
        $this->assertInternalType('array', $arrayBodyStyle);
        $this->assertCount(1, $arrayBodyStyle);
        $this->assertArrayHasKey(1, $arrayBodyStyle);
        $this->assertNull($object->getBodyStyleAtLvl(0));
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getBodyStyleAtLvl(1));
        $oParagraph = $object->getBodyStyleAtLvl(1);
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        $this->assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        $this->assertEquals((-324900 / 9525), $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(0, $oParagraph->getAlignment()->getMarginLeft());
        $this->assertEquals(32, $oParagraph->getFont()->getSize());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Style\SchemeColor', $oParagraph->getFont()->getColor());
        $this->assertEquals('tx1', $oParagraph->getFont()->getColor()->getValue());

        $arrayOtherStyle = $object->getOtherStyle();
        $this->assertInternalType('array', $arrayOtherStyle);
        $this->assertCount(1, $arrayOtherStyle);
        $this->assertArrayHasKey(0, $arrayOtherStyle);
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getOtherStyleAtLvl(0));
        $this->assertNull($object->getOtherStyleAtLvl(1));
        $oParagraph = $object->getOtherStyleAtLvl(0);
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        $this->assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        $this->assertEquals(10, $oParagraph->getFont()->getSize());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Style\SchemeColor', $oParagraph->getFont()->getColor());
        $this->assertEquals('tx1', $oParagraph->getFont()->getColor()->getValue());

        $arrayTitleStyle = $object->getTitleStyle();
        $this->assertInternalType('array', $arrayTitleStyle);
        $this->assertCount(1, $arrayTitleStyle);
        $this->assertArrayHasKey(1, $arrayTitleStyle);
        $this->assertNull($object->getTitleStyleAtLvl(0));
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getTitleStyleAtLvl(1));
        $oParagraph = $object->getTitleStyleAtLvl(1);
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        $this->assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        $this->assertEquals(44, $oParagraph->getFont()->getSize());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Style\SchemeColor', $oParagraph->getFont()->getColor());
        $this->assertEquals('lt1', $oParagraph->getFont()->getColor()->getValue());
    }

    public function testConstructDefaultFalse()
    {
        $object = new TextStyle(false);

        $this->assertInternalType('array', $object->getBodyStyle());
        $this->assertCount(0, $object->getBodyStyle());
        $this->assertInternalType('array', $object->getOtherStyle());
        $this->assertCount(0, $object->getOtherStyle());
        $this->assertInternalType('array', $object->getTitleStyle());
        $this->assertCount(0, $object->getTitleStyle());
    }

    public function testLevel()
    {
        $value = rand(0, 9);
        $object = new TextStyle(false);
        $oParagraph = new Paragraph();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, ''));
        $this->assertNull($object->getBodyStyleAtLvl(''));
        $this->assertCount(0, $object->getBodyStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, 10));
        $this->assertNull($object->getBodyStyleAtLvl(10));
        $this->assertCount(0, $object->getBodyStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, $value));
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getBodyStyleAtLvl($value));
        $this->assertCount(1, $object->getBodyStyle());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, ''));
        $this->assertNull($object->getOtherStyleAtLvl(''));
        $this->assertCount(0, $object->getOtherStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, 10));
        $this->assertNull($object->getOtherStyleAtLvl(10));
        $this->assertCount(0, $object->getOtherStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, $value));
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getOtherStyleAtLvl($value));
        $this->assertCount(1, $object->getOtherStyle());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, ''));
        $this->assertNull($object->getTitleStyleAtLvl(''));
        $this->assertCount(0, $object->getTitleStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, 10));
        $this->assertNull($object->getTitleStyleAtLvl(10));
        $this->assertCount(0, $object->getTitleStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, $value));
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getTitleStyleAtLvl($value));
        $this->assertCount(1, $object->getTitleStyle());
    }
}
