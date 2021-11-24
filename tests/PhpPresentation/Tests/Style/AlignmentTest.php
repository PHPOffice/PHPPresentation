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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Style\Alignment;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class AlignmentTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
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
     * Test get/set horizontal.
     */
    public function testSetGetHorizontal(): void
    {
        $object = new Alignment();
        $this->assertInstanceOf(Alignment::class, $object->setHorizontal(''));
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $object->getHorizontal());
        $this->assertInstanceOf(Alignment::class, $object->setHorizontal(Alignment::HORIZONTAL_GENERAL));
        $this->assertEquals(Alignment::HORIZONTAL_GENERAL, $object->getHorizontal());
    }

    /**
     * Test get/set vertical.
     */
    public function testTextDirection(): void
    {
        $object = new Alignment();
        $this->assertInstanceOf(Alignment::class, $object->setTextDirection(''));
        $this->assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
        $this->assertInstanceOf(Alignment::class, $object->setTextDirection(Alignment::TEXT_DIRECTION_VERTICAL_90));
        $this->assertEquals(Alignment::TEXT_DIRECTION_VERTICAL_90, $object->getTextDirection());
        $this->assertInstanceOf(Alignment::class, $object->setTextDirection());
        $this->assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
    }

    /**
     * Test get/set vertical.
     */
    public function testSetGetVertical(): void
    {
        $object = new Alignment();
        $this->assertInstanceOf(Alignment::class, $object->setVertical(''));
        $this->assertEquals(Alignment::VERTICAL_BASE, $object->getVertical());
        $this->assertInstanceOf(Alignment::class, $object->setVertical(Alignment::VERTICAL_AUTO));
        $this->assertEquals(Alignment::VERTICAL_AUTO, $object->getVertical());
    }

    /**
     * Test get/set min level exception.
     */
    public function testSetGetLevelExceptionMin(): void
    {
        $this->expectException(OutOfBoundsException::class);
        $this->expectExceptionMessage('The expected value (-1) is out of bounds (0, Infinite)');

        $object = new Alignment();
        $object->setLevel(-1);
    }

    /**
     * Test get/set level.
     */
    public function testSetGetLevel(): void
    {
        $object = new Alignment();
        $value = mt_rand(1, 8);
        $this->assertInstanceOf(Alignment::class, $object->setLevel($value));
        $this->assertEquals($value, $object->getLevel());
    }

    /**
     * Test get/set indent.
     */
    public function testSetGetIndent(): void
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Alignment::class, $object->setIndent($value));
        $this->assertEquals(0, $object->getIndent());
        $value = mt_rand(-100, 0);
        $this->assertInstanceOf(Alignment::class, $object->setIndent($value));
        $this->assertEquals($value, $object->getIndent());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Alignment::class, $object->setIndent($value));
        $this->assertEquals($value, $object->getIndent());
        $value = mt_rand(-100, 0);
        $this->assertInstanceOf(Alignment::class, $object->setIndent($value));
        $this->assertEquals($value, $object->getIndent());
    }

    /**
     * Test get/set margin bottom.
     */
    public function testSetGetMarginBottom(): void
    {
        $object = new Alignment();
        $value = mt_rand(0, 100);
        $this->assertInstanceOf(Alignment::class, $object->setMarginBottom($value));
        $this->assertEquals($value, $object->getMarginBottom());
        $this->assertInstanceOf(Alignment::class, $object->setMarginBottom());
        $this->assertEquals(0, $object->getMarginBottom());
    }

    /**
     * Test get/set margin left.
     */
    public function testSetGetMarginLeft(): void
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        $this->assertEquals(0, $object->getMarginLeft());
        $value = mt_rand(-100, 0);
        $this->assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        $this->assertEquals($value, $object->getMarginLeft());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        $this->assertEquals($value, $object->getMarginLeft());
        $value = mt_rand(-100, 0);
        $this->assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        $this->assertEquals($value, $object->getMarginLeft());
    }

    /**
     * Test get/set margin right.
     */
    public function testSetGetMarginRight(): void
    {
        $object = new Alignment();
        // != Alignment::HORIZONTAL_GENERAL
        $object->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        $this->assertEquals(0, $object->getMarginRight());
        $value = mt_rand(-100, 0);
        $this->assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        $this->assertEquals($value, $object->getMarginRight());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        $this->assertEquals($value, $object->getMarginRight());
        $value = mt_rand(-100, 0);
        $this->assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        $this->assertEquals($value, $object->getMarginRight());
    }

    /**
     * Test get/set margin top.
     */
    public function testSetGetMarginTop(): void
    {
        $object = new Alignment();
        $value = mt_rand(1, 100);
        $this->assertInstanceOf(Alignment::class, $object->setMarginTop($value));
        $this->assertEquals($value, $object->getMarginTop());
        $this->assertInstanceOf(Alignment::class, $object->setMarginTop());
        $this->assertEquals(0, $object->getMarginTop());
    }

    public function testRTL(): void
    {
        $object = new Alignment();
        $this->assertFalse($object->isRTL());
        $this->assertInstanceOf(Alignment::class, $object->setIsRTL(true));
        $this->assertTrue($object->isRTL());
        $this->assertInstanceOf(Alignment::class, $object->setIsRTL(false));
        $this->assertFalse($object->isRTL());
        $this->assertInstanceOf(Alignment::class, $object->setIsRTL());
        $this->assertFalse($object->isRTL());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $value = rand(1, 100);

        $object = new Alignment();
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
}
