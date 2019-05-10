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
use PhpOffice\PhpPresentation\Style\Shadow;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class ShadowTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Shadow();
        $this->assertFalse($object->isVisible());
        $this->assertEquals(6, $object->getBlurRadius());
        $this->assertEquals(2, $object->getDistance());
        $this->assertEquals(0, $object->getDirection());
        $this->assertEquals(Shadow::SHADOW_BOTTOM_RIGHT, $object->getAlignment());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        $this->assertEquals(50, $object->getAlpha());
    }

    /**
     * Test get/set alignment
     */
    public function testSetGetAlignment()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setAlignment());
        $this->assertEquals(Shadow::SHADOW_BOTTOM_RIGHT, $object->getAlignment());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setAlignment(Shadow::SHADOW_CENTER));
        $this->assertEquals(Shadow::SHADOW_CENTER, $object->getAlignment());
    }

    /**
     * Test get/set alpha
     */
    public function testSetGetAlpha()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setAlpha());
        $this->assertEquals(0, $object->getAlpha());
        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setAlpha($value));
        $this->assertEquals($value, $object->getAlpha());
    }

    /**
     * Test get/set blur radius
     */
    public function testSetGetBlurRadius()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setBlurRadius());
        $this->assertEquals(6, $object->getBlurRadius());
        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setBlurRadius($value));
        $this->assertEquals($value, $object->getBlurRadius());
    }

    /**
     * Test get/set color
     */
    public function testSetGetColor()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setColor());
        $this->assertNull($object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setColor(new Color(Color::COLOR_BLUE)));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertEquals(Color::COLOR_BLUE, $object->getColor()->getARGB());
    }

    /**
     * Test get/set direction
     */
    public function testSetGetDirection()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setDirection());
        $this->assertEquals(0, $object->getDirection());
        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setDirection($value));
        $this->assertEquals($value, $object->getDirection());
    }

    /**
     * Test get/set distance
     */
    public function testSetGetDistance()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setDistance());
        $this->assertEquals(2, $object->getDistance());
        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setDistance($value));
        $this->assertEquals($value, $object->getDistance());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Shadow();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set visible
     */
    public function testSetIsVisible()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setVisible());
        $this->assertFalse($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setVisible(false));
        $this->assertFalse($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Shadow', $object->setVisible(true));
        $this->assertTrue($object->isVisible());
    }
}
