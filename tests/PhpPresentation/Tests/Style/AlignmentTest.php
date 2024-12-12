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
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $object->getHorizontal());
        self::assertEquals(Alignment::VERTICAL_BASE, $object->getVertical());
        self::assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
        self::assertEquals(0, $object->getLevel());
        self::assertEquals(0, $object->getIndent());
        self::assertEquals(0, $object->getMarginLeft());
        self::assertEquals(0, $object->getMarginRight());
        self::assertEquals(0, $object->getMarginTop());
        self::assertEquals(0, $object->getMarginBottom());
    }

    /**
     * Test get/set horizontal.
     */
    public function testSetGetHorizontal(): void
    {
        $object = new Alignment();
        self::assertInstanceOf(Alignment::class, $object->setHorizontal(''));
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $object->getHorizontal());
        self::assertInstanceOf(Alignment::class, $object->setHorizontal(Alignment::HORIZONTAL_GENERAL));
        self::assertEquals(Alignment::HORIZONTAL_GENERAL, $object->getHorizontal());
    }

    /**
     * Test get/set vertical.
     */
    public function testTextDirection(): void
    {
        $object = new Alignment();
        self::assertInstanceOf(Alignment::class, $object->setTextDirection(''));
        self::assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
        self::assertInstanceOf(Alignment::class, $object->setTextDirection(Alignment::TEXT_DIRECTION_VERTICAL_90));
        self::assertEquals(Alignment::TEXT_DIRECTION_VERTICAL_90, $object->getTextDirection());
        self::assertInstanceOf(Alignment::class, $object->setTextDirection());
        self::assertEquals(Alignment::TEXT_DIRECTION_HORIZONTAL, $object->getTextDirection());
    }

    /**
     * Test get/set vertical.
     */
    public function testSetGetVertical(): void
    {
        $object = new Alignment();
        self::assertInstanceOf(Alignment::class, $object->setVertical(''));
        self::assertEquals(Alignment::VERTICAL_BASE, $object->getVertical());
        self::assertInstanceOf(Alignment::class, $object->setVertical(Alignment::VERTICAL_AUTO));
        self::assertEquals(Alignment::VERTICAL_AUTO, $object->getVertical());
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
        self::assertInstanceOf(Alignment::class, $object->setLevel($value));
        self::assertEquals($value, $object->getLevel());
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
        self::assertInstanceOf(Alignment::class, $object->setIndent($value));
        self::assertEquals(0, $object->getIndent());
        $value = mt_rand(-100, 0);
        self::assertInstanceOf(Alignment::class, $object->setIndent($value));
        self::assertEquals($value, $object->getIndent());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Alignment::class, $object->setIndent($value));
        self::assertEquals($value, $object->getIndent());
        $value = mt_rand(-100, 0);
        self::assertInstanceOf(Alignment::class, $object->setIndent($value));
        self::assertEquals($value, $object->getIndent());
    }

    /**
     * Test get/set margin bottom.
     */
    public function testSetGetMarginBottom(): void
    {
        $object = new Alignment();
        $value = mt_rand(0, 100);
        self::assertInstanceOf(Alignment::class, $object->setMarginBottom($value));
        self::assertEquals($value, $object->getMarginBottom());
        self::assertInstanceOf(Alignment::class, $object->setMarginBottom());
        self::assertEquals(0, $object->getMarginBottom());
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
        self::assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        self::assertEquals(0, $object->getMarginLeft());
        $value = mt_rand(-100, 0);
        self::assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        self::assertEquals($value, $object->getMarginLeft());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        self::assertEquals($value, $object->getMarginLeft());
        $value = mt_rand(-100, 0);
        self::assertInstanceOf(Alignment::class, $object->setMarginLeft($value));
        self::assertEquals($value, $object->getMarginLeft());
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
        self::assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        self::assertEquals(0, $object->getMarginRight());
        $value = mt_rand(-100, 0);
        self::assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        self::assertEquals($value, $object->getMarginRight());

        $object->setHorizontal(Alignment::HORIZONTAL_GENERAL);
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        self::assertEquals($value, $object->getMarginRight());
        $value = mt_rand(-100, 0);
        self::assertInstanceOf(Alignment::class, $object->setMarginRight($value));
        self::assertEquals($value, $object->getMarginRight());
    }

    /**
     * Test get/set margin top.
     */
    public function testSetGetMarginTop(): void
    {
        $object = new Alignment();
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Alignment::class, $object->setMarginTop($value));
        self::assertEquals($value, $object->getMarginTop());
        self::assertInstanceOf(Alignment::class, $object->setMarginTop());
        self::assertEquals(0, $object->getMarginTop());
    }

    public function testRTL(): void
    {
        $object = new Alignment();
        self::assertFalse($object->isRTL());
        self::assertInstanceOf(Alignment::class, $object->setIsRTL(true));
        self::assertTrue($object->isRTL());
        self::assertInstanceOf(Alignment::class, $object->setIsRTL(false));
        self::assertFalse($object->isRTL());
        self::assertInstanceOf(Alignment::class, $object->setIsRTL());
        self::assertFalse($object->isRTL());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $value = mt_rand(1, 100);

        $object = new Alignment();
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }
}
