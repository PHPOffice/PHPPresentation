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

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class ColorTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Color();
        $this->assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $object = new Color(COLOR::COLOR_BLUE);
        $this->assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }

    /**
     * Test Alpha
     */
    public function testAlpha()
    {
        $object = new Color();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB());
        $this->assertEquals(100, $object->getAlpha());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(Color::COLOR_BLUE));
        $this->assertEquals(100, $object->getAlpha());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB('AA0000FF'));
        $this->assertEquals(66.67, $object->getAlpha());
    }

    /**
     * Test get/set ARGB
     */
    public function testSetGetARGB()
    {
        $object = new Color();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB());
        $this->assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(''));
        $this->assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setARGB(Color::COLOR_BLUE));
        $this->assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }

    /**
     * Test get/set RGB
     */
    public function testSetGetRGB()
    {
        $object = new Color();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB());
        $this->assertEquals('000000', $object->getRGB());
        $this->assertEquals('FF000000', $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB(''));
        $this->assertEquals('000000', $object->getRGB());
        $this->assertEquals('FF000000', $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB('555'));
        $this->assertEquals('555', $object->getRGB());
        $this->assertEquals('FF555', $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->setRGB('6666'));
        $this->assertEquals('FF6666', $object->getRGB());
        $this->assertEquals('FF6666', $object->getARGB());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Color();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
}
