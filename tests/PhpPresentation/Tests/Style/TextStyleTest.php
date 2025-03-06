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

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\SchemeColor;
use PhpOffice\PhpPresentation\Style\TextStyle;
use PHPUnit\Framework\TestCase;

class TextStyleTest extends TestCase
{
    public function testConstructDefaultTrue(): void
    {
        /** @var TextStyle $object */
        $object = new TextStyle();

        $arrayBodyStyle = $object->getBodyStyle();
        self::assertIsArray($arrayBodyStyle);
        self::assertCount(1, $arrayBodyStyle);
        self::assertArrayHasKey(1, $arrayBodyStyle);
        self::assertNull($object->getBodyStyleAtLvl(0));
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getBodyStyleAtLvl(1));
        $oParagraph = $object->getBodyStyleAtLvl(1);
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        self::assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        self::assertEquals((-342900 / 9525), $oParagraph->getAlignment()->getIndent());
        self::assertEquals(0, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(32, $oParagraph->getFont()->getSize());
        /** @var SchemeColor $color */
        $color = $oParagraph->getFont()->getColor();
        self::assertInstanceOf(SchemeColor::class, $color);
        self::assertEquals('tx1', $color->getValue());

        $arrayOtherStyle = $object->getOtherStyle();
        self::assertIsArray($arrayOtherStyle);
        self::assertCount(1, $arrayOtherStyle);
        self::assertArrayHasKey(0, $arrayOtherStyle);
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getOtherStyleAtLvl(0));
        self::assertNull($object->getOtherStyleAtLvl(1));
        $oParagraph = $object->getOtherStyleAtLvl(0);
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        self::assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        self::assertEquals(10, $oParagraph->getFont()->getSize());
        /** @var SchemeColor $color */
        $color = $oParagraph->getFont()->getColor();
        self::assertInstanceOf(SchemeColor::class, $color);
        self::assertEquals('tx1', $color->getValue());

        $arrayTitleStyle = $object->getTitleStyle();
        self::assertIsArray($arrayTitleStyle);
        self::assertCount(1, $arrayTitleStyle);
        self::assertArrayHasKey(1, $arrayTitleStyle);
        self::assertNull($object->getTitleStyleAtLvl(0));
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getTitleStyleAtLvl(1));
        $oParagraph = $object->getTitleStyleAtLvl(1);
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $oParagraph);
        self::assertEquals(Alignment::HORIZONTAL_CENTER, $oParagraph->getAlignment()->getHorizontal());
        self::assertEquals(44, $oParagraph->getFont()->getSize());
        /** @var SchemeColor $color */
        $color = $oParagraph->getFont()->getColor();
        self::assertInstanceOf(SchemeColor::class, $color);
        self::assertEquals('lt1', $color->getValue());
    }

    public function testConstructDefaultFalse(): void
    {
        $object = new TextStyle(false);

        self::assertIsArray($object->getBodyStyle());
        self::assertCount(0, $object->getBodyStyle());
        self::assertIsArray($object->getOtherStyle());
        self::assertCount(0, $object->getOtherStyle());
        self::assertIsArray($object->getTitleStyle());
        self::assertCount(0, $object->getTitleStyle());
    }

    public function testLevel(): void
    {
        $value = mt_rand(0, 9);
        $object = new TextStyle(false);
        $oParagraph = new Paragraph();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, null));
        self::assertNull($object->getBodyStyleAtLvl(null));
        self::assertCount(0, $object->getBodyStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, 10));
        self::assertNull($object->getBodyStyleAtLvl(10));
        self::assertCount(0, $object->getBodyStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setBodyStyleAtLvl($oParagraph, $value));
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getBodyStyleAtLvl($value));
        self::assertCount(1, $object->getBodyStyle());

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, null));
        self::assertNull($object->getOtherStyleAtLvl(null));
        self::assertCount(0, $object->getOtherStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, 10));
        self::assertNull($object->getOtherStyleAtLvl(10));
        self::assertCount(0, $object->getOtherStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setOtherStyleAtLvl($oParagraph, $value));
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getOtherStyleAtLvl($value));
        self::assertCount(1, $object->getOtherStyle());

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, null));
        self::assertNull($object->getTitleStyleAtLvl(null));
        self::assertCount(0, $object->getTitleStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, 10));
        self::assertNull($object->getTitleStyleAtLvl(10));
        self::assertCount(0, $object->getTitleStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->setTitleStyleAtLvl($oParagraph, $value));
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\RichText\Paragraph', $object->getTitleStyleAtLvl($value));
        self::assertCount(1, $object->getTitleStyle());
    }
}
