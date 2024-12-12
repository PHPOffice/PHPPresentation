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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\DocumentLayout;
use PHPUnit\Framework\TestCase;

/**
 * Test class for DocumentLayout.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\DocumentLayout
 */
class DocumentLayoutTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new DocumentLayout();

        self::assertEquals('screen4x3', $object->getDocumentLayout());
        self::assertEquals(9144000, $object->getCX());
        self::assertEquals(6858000, $object->getCY());
    }

    /**
     * Test set custom layout.
     */
    public function testSetCustomLayout(): void
    {
        $object = new DocumentLayout();
        $object->setDocumentLayout(['cx' => 6858000, 'cy' => 9144000], false);
        self::assertEquals(DocumentLayout::LAYOUT_CUSTOM, $object->getDocumentLayout());
        self::assertEquals(9144000, $object->getCX());
        self::assertEquals(6858000, $object->getCY());
        $object->setDocumentLayout(['cx' => 6858000, 'cy' => 9144000], true);
        self::assertEquals(DocumentLayout::LAYOUT_CUSTOM, $object->getDocumentLayout());
        self::assertEquals(6858000, $object->getCX());
        self::assertEquals(9144000, $object->getCY());
    }

    /**
     * Test set custom layout.
     */
    public function testSetCustomLayoutWithString(): void
    {
        $object = new DocumentLayout();
        $object->setDocumentLayout(DocumentLayout::LAYOUT_CUSTOM);
        self::assertEquals(DocumentLayout::LAYOUT_CUSTOM, $object->getDocumentLayout());
        // Default value
        self::assertEquals(9144000, $object->getCX());
        self::assertEquals(6858000, $object->getCY());

        $object->setCX(13.333, DocumentLayout::UNIT_CENTIMETER);
        $object->setCY(7.5, DocumentLayout::UNIT_CENTIMETER);
        self::assertEquals(4799880, $object->getCX());
        self::assertEquals(2700000, $object->getCY());
    }

    public function testCX(): void
    {
        $value = mt_rand(1, 100000);
        $object = new DocumentLayout();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value));
        self::assertEquals($value, $object->getCX());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_CENTIMETER));
        self::assertEquals($value, $object->getCX(DocumentLayout::UNIT_CENTIMETER));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_EMU));
        self::assertEquals($value, $object->getCX(DocumentLayout::UNIT_EMU));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_INCH));
        self::assertEquals($value, $object->getCX(DocumentLayout::UNIT_INCH));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_MILLIMETER));
        self::assertEquals($value, $object->getCX(DocumentLayout::UNIT_MILLIMETER));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_POINT));
        self::assertEquals($value, $object->getCX(DocumentLayout::UNIT_POINT));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_PIXEL));
        self::assertEquals($value, $object->getCX(DocumentLayout::UNIT_PIXEL));
    }

    public function testCY(): void
    {
        $value = mt_rand(1, 100000);
        $object = new DocumentLayout();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value));
        self::assertEquals($value, $object->getCY());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_CENTIMETER));
        self::assertEquals($value, $object->getCY(DocumentLayout::UNIT_CENTIMETER));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_EMU));
        self::assertEquals($value, $object->getCY(DocumentLayout::UNIT_EMU));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_INCH));
        self::assertEquals($value, $object->getCY(DocumentLayout::UNIT_INCH));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_MILLIMETER));
        self::assertEquals($value, $object->getCY(DocumentLayout::UNIT_MILLIMETER));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_POINT));
        self::assertEquals($value, $object->getCY(DocumentLayout::UNIT_POINT));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_PIXEL));
        self::assertEquals($value, $object->getCY(DocumentLayout::UNIT_PIXEL));
    }
}
