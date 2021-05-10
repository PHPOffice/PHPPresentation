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

use PhpOffice\PhpPresentation\Style\Color;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class ColorTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Color();
        static::assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $object = new Color(Color::COLOR_BLUE);
        static::assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }

    /**
     * Test Alpha
     */
    public function testAlpha()
    {
        $randAlpha = mt_rand(0, 100);
        $object = new Color();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB());
        static::assertEquals(100, $object->getAlpha());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB('AA0000FF'));
        static::assertEquals(66.67, $object->getAlpha());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(Color::COLOR_BLUE));
        static::assertEquals(100, $object->getAlpha());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setAlpha($randAlpha));
        static::assertEquals($randAlpha, round($object->getAlpha()));
    }

    /**
     * Test get/set ARGB
     */
    public function testSetGetARGB()
    {
        $object = new Color();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB());
        static::assertEquals(Color::COLOR_BLACK, $object->getARGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(''));
        static::assertEquals(Color::COLOR_BLACK, $object->getARGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(Color::COLOR_BLUE));
        static::assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }

    /**
     * Test get/set RGB
     */
    public function testSetGetRGB()
    {
        $object = new Color();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB());
        static::assertEquals('000000', $object->getRGB());
        static::assertEquals('FF000000', $object->getARGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB(''));
        static::assertEquals('000000', $object->getRGB());
        static::assertEquals('FF000000', $object->getARGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB('555'));
        static::assertEquals('555', $object->getRGB());
        static::assertEquals('FF555', $object->getARGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB('6666'));
        static::assertEquals('FF6666', $object->getRGB());
        static::assertEquals('FF6666', $object->getARGB());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Color();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        static::assertEquals($value, $object->getHashIndex());
    }
}
