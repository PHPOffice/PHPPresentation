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

use PhpOffice\PhpPresentation\Style\Alignment;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class AlignmentTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Alignment();
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $object->getHorizontal());
        $this->assertEquals(Alignment::VERTICAL_BASE, $object->getVertical());
        $this->assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
        $this->assertEquals(0, $object->getLevel());
        $this->assertEquals(0, $object->getIndent());
        $this->assertEquals(0, $object->getMarginLeft());
        $this->assertEquals(0, $object->getMarginRight());
        $this->assertEquals(0, $object->getMarginTop());
        $this->assertEquals(0, $object->getMarginBottom());
    }

    /**
     * Test get/set horizontal
     */
    public function testSetGetHorizontal()
    {
        $object = new Alignment();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setHorizontal(''));
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $object->getHorizontal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setHorizontal(Alignment::HORIZONTAL_GENERAL));
        $this->assertEquals(Alignment::HORIZONTAL_GENERAL, $object->getHorizontal());
    }

    /**
     * Test get/set vertical
     */
    public function testTextDirection()
    {
        $object = new Alignment();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setTextDirection(null));
        $this->assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setTextDirection(Alignment::TEXT_DIRECTION_VERTICAL_90));
        $this->assertEquals(Alignment::TEXT_DIRECTION_VERTICAL_90, $object->getTextDirection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setTextDirection());
        $this->assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
    }

    /**
     * Test get/set vertical
     */
    public function testSetGetVertical()
    {
        $object = new Alignment();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setVertical(''));
        $this->assertEquals(Alignment::VERTICAL_BASE, $object->getVertical());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setVertical(Alignment::VERTICAL_AUTO));
        $this->assertEquals(Alignment::VERTICAL_AUTO, $object->getVertical());
    }

    /**
     * Test get/set min level exception
     */
    public function testSetGetLevelExceptionMin()
    {
        $object = new Alignment();
        $this->setExpectedException('\Exception', 'Invalid value should be more than 0.');
        $object->setLevel(-1);
    }

    /**
     * Test get/set level
     */
    public function testSetGetLevel()
    {
        $object = new Alignment();
        $value = rand(1, 8);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setLevel($value));
        $this->assertEquals($value, $object->getLevel());
    }

    /**
     * Test get/set indent
     */
    public function testSetGetIndent()
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        $this->assertEquals(0, $object->getIndent());
        $value = rand(-100, 0);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        $this->assertEquals($value, $object->getIndent());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        $this->assertEquals($value, $object->getIndent());
        $value = rand(-100, 0);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        $this->assertEquals($value, $object->getIndent());
    }

    /**
     * Test get/set margin bottom
     */
    public function testSetGetMarginBottom()
    {
        $object = new Alignment();
        $value = rand(0, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginBottom($value));
        $this->assertEquals($value, $object->getMarginBottom());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginBottom());
        $this->assertEquals(0, $object->getMarginBottom());
    }

    /**
     * Test get/set margin left
     */
    public function testSetGetMarginLeft()
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        $this->assertEquals(0, $object->getMarginLeft());
        $value = rand(-100, 0);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        $this->assertEquals($value, $object->getMarginLeft());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        $this->assertEquals($value, $object->getMarginLeft());
        $value = rand(-100, 0);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        $this->assertEquals($value, $object->getMarginLeft());
    }

    /**
     * Test get/set margin right
     */
    public function testSetGetMarginRight()
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        $this->assertEquals(0, $object->getMarginRight());
        $value = rand(-100, 0);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        $this->assertEquals($value, $object->getMarginRight());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        $this->assertEquals($value, $object->getMarginRight());
        $value = rand(-100, 0);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        $this->assertEquals($value, $object->getMarginRight());
    }

    /**
     * Test get/set margin top
     */
    public function testSetGetMarginTop()
    {
        $object = new Alignment();
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginTop($value));
        $this->assertEquals($value, $object->getMarginTop());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginTop());
        $this->assertEquals(0, $object->getMarginTop());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $value = md5(rand(1, 100));

        $object = new Alignment();
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
}
