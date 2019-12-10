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
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class AlignmentTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Alignment();
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $object->getHorizontal());
        static::assertEquals(Alignment::VERTICAL_BASE, $object->getVertical());
        static::assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
        static::assertEquals(0, $object->getLevel());
        static::assertEquals(0, $object->getIndent());
        static::assertEquals(0, $object->getMarginLeft());
        static::assertEquals(0, $object->getMarginRight());
        static::assertEquals(0, $object->getMarginTop());
        static::assertEquals(0, $object->getMarginBottom());
    }

    /**
     * Test get/set horizontal
     */
    public function testSetGetHorizontal()
    {
        $object = new Alignment();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setHorizontal(''));
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $object->getHorizontal());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setHorizontal(Alignment::HORIZONTAL_GENERAL));
        static::assertEquals(Alignment::HORIZONTAL_GENERAL, $object->getHorizontal());
    }

    /**
     * Test get/set vertical
     */
    public function testTextDirection()
    {
        $object = new Alignment();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setTextDirection(null));
        static::assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setTextDirection(Alignment::TEXT_DIRECTION_VERTICAL_90));
        static::assertEquals(Alignment::TEXT_DIRECTION_VERTICAL_90, $object->getTextDirection());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setTextDirection());
        static::assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
    }

    /**
     * Test get/set vertical
     */
    public function testSetGetVertical()
    {
        $object = new Alignment();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setVertical(''));
        static::assertEquals(Alignment::VERTICAL_BASE, $object->getVertical());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setVertical(Alignment::VERTICAL_AUTO));
        static::assertEquals(Alignment::VERTICAL_AUTO, $object->getVertical());
    }

    /**
     * Test get/set min level exception
     */
    public function testSetGetLevelExceptionMin()
    {
        $object = new Alignment();
        if (method_exists($this, 'setExpectedException')) {
            $this->setExpectedException('\Exception', 'Invalid value should be more than 0.');
        }
        if (method_exists($this, 'expectException')) {
            $this->expectException('\Exception', 'Invalid value should be more than 0.');
        }
        $object->setLevel(-1);
    }

    /**
     * Test get/set level
     */
    public function testSetGetLevel()
    {
        $object = new Alignment();
        $value = mt_rand(1, 8);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setLevel($value));
        static::assertEquals($value, $object->getLevel());
    }

    /**
     * Test get/set indent
     */
    public function testSetGetIndent()
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        static::assertEquals(0, $object->getIndent());
        $value = mt_rand(-100, 0);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        static::assertEquals($value, $object->getIndent());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        static::assertEquals($value, $object->getIndent());
        $value = mt_rand(-100, 0);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setIndent($value));
        static::assertEquals($value, $object->getIndent());
    }

    /**
     * Test get/set margin bottom
     */
    public function testSetGetMarginBottom()
    {
        $object = new Alignment();
        $value = mt_rand(0, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginBottom($value));
        static::assertEquals($value, $object->getMarginBottom());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginBottom());
        static::assertEquals(0, $object->getMarginBottom());
    }

    /**
     * Test get/set margin left
     */
    public function testSetGetMarginLeft()
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        static::assertEquals(0, $object->getMarginLeft());
        $value = mt_rand(-100, 0);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        static::assertEquals($value, $object->getMarginLeft());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        static::assertEquals($value, $object->getMarginLeft());
        $value = mt_rand(-100, 0);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginLeft($value));
        static::assertEquals($value, $object->getMarginLeft());
    }

    /**
     * Test get/set margin right
     */
    public function testSetGetMarginRight()
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        static::assertEquals(0, $object->getMarginRight());
        $value = mt_rand(-100, 0);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        static::assertEquals($value, $object->getMarginRight());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        static::assertEquals($value, $object->getMarginRight());
        $value = mt_rand(-100, 0);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginRight($value));
        static::assertEquals($value, $object->getMarginRight());
    }

    /**
     * Test get/set margin top
     */
    public function testSetGetMarginTop()
    {
        $object = new Alignment();
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginTop($value));
        static::assertEquals($value, $object->getMarginTop());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->setMarginTop());
        static::assertEquals(0, $object->getMarginTop());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $value = md5(rand(1, 100));

        $object = new Alignment();
        $object->setHashIndex($value);
        static::assertEquals($value, $object->getHashIndex());
    }
}
