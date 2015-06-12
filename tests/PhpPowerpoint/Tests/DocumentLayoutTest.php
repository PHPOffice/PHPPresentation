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

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\DocumentLayout;

/**
 * Test class for DocumentLayout
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\DocumentLayout
 */
class DocumentLayoutTest extends \PHPUnit_Framework_TestCase
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
        $this->assertEquals(9144000 / 36000, $object->getLayoutXmilli());
        $this->assertEquals(6858000 / 36000, $object->getLayoutYmilli());
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
        $value = rand(1, 100000);
        $object = new DocumentLayout();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCX($value));
        $this->assertEquals($value, $object->getCX());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_CENTIMETER));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_CENTIMETER));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_EMU));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_EMU));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_INCH));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_INCH));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getLayoutXmilli());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_POINT));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_POINT));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCX($value, DocumentLayout::UNIT_PIXEL));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_PIXEL));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setLayoutXmilli($value));
        $this->assertEquals($value, $object->getCX(DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getLayoutXmilli());
    }

    public function testCY()
    {
        $value = rand(1, 100000);
        $object = new DocumentLayout();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCY($value));
        $this->assertEquals($value, $object->getCY());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_CENTIMETER));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_CENTIMETER));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_EMU));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_EMU));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_INCH));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_INCH));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getLayoutYmilli());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_POINT));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_POINT));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setCY($value, DocumentLayout::UNIT_PIXEL));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_PIXEL));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\DocumentLayout', $object->setLayoutYmilli($value));
        $this->assertEquals($value, $object->getCY(DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals($value, $object->getLayoutYmilli());
    }
}
