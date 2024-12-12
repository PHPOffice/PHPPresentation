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

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Shadow;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class ShadowTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new Shadow();
        self::assertFalse($object->isVisible());
        self::assertEquals(6, $object->getBlurRadius());
        self::assertEquals(2, $object->getDistance());
        self::assertEquals(0, $object->getDirection());
        self::assertEquals(Shadow::SHADOW_BOTTOM_RIGHT, $object->getAlignment());
        self::assertInstanceOf(Color::class, $object->getColor());
        self::assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        self::assertEquals(50, $object->getAlpha());
    }

    /**
     * Test get/set alignment.
     */
    public function testSetGetAlignment(): void
    {
        $object = new Shadow();
        self::assertInstanceOf(Shadow::class, $object->setAlignment());
        self::assertEquals(Shadow::SHADOW_BOTTOM_RIGHT, $object->getAlignment());
        self::assertInstanceOf(Shadow::class, $object->setAlignment(Shadow::SHADOW_CENTER));
        self::assertEquals(Shadow::SHADOW_CENTER, $object->getAlignment());
    }

    /**
     * Test get/set alpha.
     */
    public function testSetGetAlpha(): void
    {
        $object = new Shadow();
        self::assertInstanceOf(Shadow::class, $object->setAlpha());
        self::assertEquals(0, $object->getAlpha());
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Shadow::class, $object->setAlpha($value));
        self::assertEquals($value, $object->getAlpha());
    }

    /**
     * Test get/set blur radius.
     */
    public function testSetGetBlurRadius(): void
    {
        $object = new Shadow();
        self::assertInstanceOf(Shadow::class, $object->setBlurRadius());
        self::assertEquals(6, $object->getBlurRadius());
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Shadow::class, $object->setBlurRadius($value));
        self::assertEquals($value, $object->getBlurRadius());
    }

    /**
     * Test get/set color.
     */
    public function testSetGetColor(): void
    {
        $object = new Shadow();
        self::assertInstanceOf(Shadow::class, $object->setColor());
        self::assertNull($object->getColor());
        self::assertInstanceOf(Shadow::class, $object->setColor(new Color(Color::COLOR_BLUE)));
        self::assertInstanceOf(Color::class, $object->getColor());
        self::assertEquals(Color::COLOR_BLUE, $object->getColor()->getARGB());
    }

    /**
     * Test get/set direction.
     */
    public function testSetGetDirection(): void
    {
        $object = new Shadow();
        self::assertInstanceOf(Shadow::class, $object->setDirection());
        self::assertEquals(0, $object->getDirection());
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Shadow::class, $object->setDirection($value));
        self::assertEquals($value, $object->getDirection());
    }

    /**
     * Test get/set distance.
     */
    public function testSetGetDistance(): void
    {
        $object = new Shadow();
        self::assertInstanceOf(Shadow::class, $object->setDistance());
        self::assertEquals(2, $object->getDistance());
        $value = mt_rand(1, 100);
        self::assertInstanceOf(Shadow::class, $object->setDistance($value));
        self::assertEquals($value, $object->getDistance());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Shadow();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set visible.
     */
    public function testSetIsVisible(): void
    {
        $object = new Shadow();
        self::assertInstanceOf(Shadow::class, $object->setVisible());
        self::assertFalse($object->isVisible());
        self::assertInstanceOf(Shadow::class, $object->setVisible(false));
        self::assertFalse($object->isVisible());
        self::assertInstanceOf(Shadow::class, $object->setVisible(true));
        self::assertTrue($object->isVisible());
    }
}
