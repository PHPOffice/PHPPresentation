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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\DocumentLayout;
use PHPUnit\Framework\TestCase;

/**
 * Test class for DocumentLayout
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\DocumentLayout
 */
class DocumentLayoutTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new DocumentLayout();

        $this->assertEquals('screen4x3', $object->getDocumentLayout());
        $this->assertEquals(9144000, $object->getCX());
        $this->assertEquals(6858000, $object->getCY());
    }

    /**
     * Test set custom layout
     */
    public function testSetCustomLayout()
    {
        $object = new DocumentLayout();
        $object->setDocumentLayout(array('cx' => 6858000, 'cy' => 9144000), false);
        $this->assertEquals(DocumentLayout::LAYOUT_CUSTOM, $object->getDocumentLayout());
        $this->assertEquals(9144000, $object->getCX());
        $this->assertEquals(6858000, $object->getCY());
        $object->setDocumentLayout(array('cx' => 6858000, 'cy' => 9144000), true);
        $this->assertEquals(DocumentLayout::LAYOUT_CUSTOM, $object->getDocumentLayout());
        $this->assertEquals(6858000, $object->getCX());
        $this->assertEquals(9144000, $object->getCY());
    }

    public function testCX()
    {
        $value = mt_rand(1, 100000);
        $object = new DocumentLayout();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value));
        $this->assertEquals($value, $object->getCX());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_CENTIMETER));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_CENTIMETER));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_EMU));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_EMU));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_INCH));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_INCH));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_MILLIMETER));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_POINT));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_POINT));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_PIXEL));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_PIXEL));
    }

    public function testCY()
    {
        $value = mt_rand(1, 100000);
        $object = new DocumentLayout();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value));
        $this->assertEquals($value, $object->getCY());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_CENTIMETER));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_CENTIMETER));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_EMU));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_EMU));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_INCH));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_INCH));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_MILLIMETER));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_POINT));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_POINT));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_PIXEL));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_PIXEL));
    }
}
