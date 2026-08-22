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
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class ColorTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new Color();
        self::assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $object = new Color(Color::COLOR_BLUE);
        self::assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }

    /**
     * Test Alpha.
     */
    public function testAlpha(): void
    {
        $randAlpha = mt_rand(0, 100);
        $object = new Color();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB());
        self::assertEquals(100, $object->getAlpha());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB('AA0000FF'));
        self::assertEquals(67, $object->getAlpha());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(Color::COLOR_BLUE));
        self::assertEquals(100, $object->getAlpha());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setAlpha($randAlpha));
        self::assertEquals($randAlpha, round($object->getAlpha()));
    }

    /**
     * Test Alpha of a colour written as a plain RGB.
     */
    public function testAlphaOfAnRGB(): void
    {
        $object = new Color('D0D0D0');
        self::assertEquals(100, $object->getAlpha());
        self::assertEquals('D0D0D0', $object->getRGB());

        // The alpha goes in front of the colour, and does not eat the first two characters of it
        $object->setAlpha(50);
        self::assertEquals('80D0D0D0', $object->getARGB());
        self::assertEquals('D0D0D0', $object->getRGB());
        self::assertEquals(50, $object->getAlpha());
    }

    /**
     * Test get/set ARGB.
     */
    public function testSetGetARGB(): void
    {
        $object = new Color();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB());
        self::assertEquals(Color::COLOR_BLACK, $object->getARGB());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(''));
        self::assertEquals(Color::COLOR_BLACK, $object->getARGB());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(Color::COLOR_BLUE));
        self::assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }

    /**
     * Test get/set RGB.
     */
    public function testSetGetRGB(): void
    {
        $object = new Color();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB());
        self::assertEquals('000000', $object->getRGB());
        self::assertEquals('FF000000', $object->getARGB());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB(''));
        self::assertEquals('000000', $object->getRGB());
        self::assertEquals('FF000000', $object->getARGB());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB('555'));
        self::assertEquals('555', $object->getRGB());
        self::assertEquals('FF555', $object->getARGB());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB('6666'));
        self::assertEquals('FF6666', $object->getRGB());
        self::assertEquals('FF6666', $object->getARGB());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Color();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }
}
