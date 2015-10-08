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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\Drawing;

/**
 * Test class for Drawing element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Drawing
 */
class DrawingTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Drawing();

        $this->assertEmpty($object->getPath());
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage File  not found!
     */
    public function testPathBasic()
    {
        $object = new Drawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setPath());
    }

    public function testPathWithoutVerifyFile()
    {
        $object = new Drawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setPath('', false));
        $this->assertEmpty($object->getPath());
    }

    public function testPathWithRealFile()
    {
        $object = new Drawing();

        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setPath($imagePath, false));
        $this->assertEquals($imagePath, $object->getPath());
        $this->assertEquals(0, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
    }

    public function testGetSetDescription()
    {
        $object = new Drawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setDescription());
        $this->assertEmpty($object->getDescription());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setDescription('TEST'));
        $this->assertEquals('TEST', $object->getDescription());
    }

    public function testGetSetResizeProportional()
    {
        $object = new Drawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setResizeProportional());
        $this->assertTrue($object->isResizeProportional());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setResizeProportional(false));
        $this->assertFalse($object->isResizeProportional());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setResizeProportional(true));
        $this->assertTrue($object->isResizeProportional());
    }

    public function testGetSetHeightWidth()
    {
        $object = new Drawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setHeight());
        $this->assertEquals(0, $object->getHeight());
    }
    /**
     * Value : Resize Proportional (false)
     */
    public function testGetSetHeightWidthResizeProportionalFalse()
    {
        // Valeur : Resize Proportional (false)
        $object = new Drawing();
        $value = rand(0, 100);
        $object->setResizeProportional(false);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }

    /**
     * Value : Resize Proportional (true)
     */
    public function testGetSetHeightWidthResizeProportionalTrue()
    {
        $object = new Drawing();
        $valueWidth = rand(1, 100);
        $valueHeight = rand(1, 100);
        $object->setResizeProportional(false);
        $object->setWidth($valueWidth);
        $object->setHeight($valueHeight);
        $object->setResizeProportional(true);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setWidth($valueWidth));
        $this->assertEquals($valueWidth, $object->getWidth());
        $this->assertEquals(round($valueWidth * ($valueHeight / $valueWidth)), $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setHeight($valueHeight));
        $this->assertEquals($valueHeight, $object->getHeight());
        $this->assertEquals(round($valueHeight * ($valueWidth / $valueHeight)), $object->getWidth());
    }

    public function testSetWidthAndHeight()
    {
        $object = new Drawing();
        $valueWidth = rand(1, 100);
        $valueHeight = $valueWidth / 2;
        $object->setResizeProportional(false);
        $object->setWidth($valueWidth);
        $object->setHeight($valueHeight);
        $object->setResizeProportional(true);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setWidthAndHeight($valueHeight, $valueWidth));
        $this->assertEquals($valueHeight, $object->getWidth());
        $this->assertEquals(ceil($valueHeight * ($valueHeight/$valueWidth)), $object->getHeight());

        $object = new Drawing();
        $valueWidth = rand(1, 100);
        $valueHeight = $valueWidth / 2;
        $object->setResizeProportional(false);
        $object->setWidth($valueWidth);
        $object->setHeight($valueHeight);
        $object->setResizeProportional(true);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object->setWidthAndHeight($valueWidth, $valueHeight));
        $this->assertEquals($valueHeight, $object->getHeight());
        $this->assertEquals(ceil($valueWidth * ($valueHeight/$valueHeight)), $object->getWidth());
    }
}
