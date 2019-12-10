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
use PHPUnit\Framework\TestCase;

class TextStyleTest extends TestCase
{
    public function testConstructDefaultTrue()
    {
        /**
         * @var \PhpOffice\PhpPresentation\Shape\RichText\Paragraph $oParagraph
         */
        $object = new TextStyle();

        $arrayBodyStyle = $object->getBodyStyle();
        static::assertInternalType('array', $arrayBodyStyle);
        static::assertCount(1, $arrayBodyStyle);
        static::assertArrayHasKey(1, $arrayBodyStyle);
        static::assertNull($object->getBodyStyleAtLvl(0));
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getBodyStyleAtLvl(1));
        $oParagraph = $object->getBodyStyleAtLvl(1);
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        static::assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals((-324900 / 9525), $oParagraph->getAlignment()->getIndent());
        static::assertEquals(0, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(32, $oParagraph->getFont()->getSize());
        static::assertInstanceOf('PhpOffice\PhpPresentation\Style\SchemeColor', $oParagraph->getFont()->getColor());
        static::assertEquals('tx1', $oParagraph->getFont()->getColor()->getValue());

        $arrayOtherStyle = $object->getOtherStyle();
        static::assertInternalType('array', $arrayOtherStyle);
        static::assertCount(1, $arrayOtherStyle);
        static::assertArrayHasKey(0, $arrayOtherStyle);
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getOtherStyleAtLvl(0));
        static::assertNull($object->getOtherStyleAtLvl(1));
        $oParagraph = $object->getOtherStyleAtLvl(0);
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        static::assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(10, $oParagraph->getFont()->getSize());
        static::assertInstanceOf('PhpOffice\PhpPresentation\Style\SchemeColor', $oParagraph->getFont()->getColor());
        static::assertEquals('tx1', $oParagraph->getFont()->getColor()->getValue());

        $arrayTitleStyle = $object->getTitleStyle();
        static::assertInternalType('array', $arrayTitleStyle);
        static::assertCount(1, $arrayTitleStyle);
        static::assertArrayHasKey(1, $arrayTitleStyle);
        static::assertNull($object->getTitleStyleAtLvl(0));
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getTitleStyleAtLvl(1));
        $oParagraph = $object->getTitleStyleAtLvl(1);
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        static::assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(44, $oParagraph->getFont()->getSize());
        static::assertInstanceOf('PhpOffice\PhpPresentation\Style\SchemeColor', $oParagraph->getFont()->getColor());
        static::assertEquals('lt1', $oParagraph->getFont()->getColor()->getValue());
    }

    public function testConstructDefaultFalse()
    {
        $object = new TextStyle(false);

        static::assertInternalType('array', $object->getBodyStyle());
        static::assertCount(0, $object->getBodyStyle());
        static::assertInternalType('array', $object->getOtherStyle());
        static::assertCount(0, $object->getOtherStyle());
        static::assertInternalType('array', $object->getTitleStyle());
        static::assertCount(0, $object->getTitleStyle());
    }

    public function testLevel()
    {
        $value = mt_rand(0, 9);
        $object = new TextStyle(false);
        $oParagraph = new Paragraph();

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, ''));
        static::assertNull($object->getBodyStyleAtLvl(''));
        static::assertCount(0, $object->getBodyStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, 10));
        static::assertNull($object->getBodyStyleAtLvl(10));
        static::assertCount(0, $object->getBodyStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, $value));
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getBodyStyleAtLvl($value));
        static::assertCount(1, $object->getBodyStyle());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, ''));
        static::assertNull($object->getOtherStyleAtLvl(''));
        static::assertCount(0, $object->getOtherStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, 10));
        static::assertNull($object->getOtherStyleAtLvl(10));
        static::assertCount(0, $object->getOtherStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, $value));
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getOtherStyleAtLvl($value));
        static::assertCount(1, $object->getOtherStyle());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, ''));
        static::assertNull($object->getTitleStyleAtLvl(''));
        static::assertCount(0, $object->getTitleStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, 10));
        static::assertNull($object->getTitleStyleAtLvl(10));
        static::assertCount(0, $object->getTitleStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, $value));
        static::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getTitleStyleAtLvl($value));
        static::assertCount(1, $object->getTitleStyle());
    }
}
