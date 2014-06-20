<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Shadow;
use PhpOffice\PhpPowerpoint\Style\Alignment;

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\PhpPowerpoint
 */
class ShadowTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct ()
    {
        $object = new Shadow();
        $this->assertFalse($object->isVisible());
        $this->assertEquals(6, $object->getBlurRadius());
        $this->assertEquals(2, $object->getDistance());
        $this->assertEquals(0, $object->getDirection());
        $this->assertEquals(Shadow::SHADOW_BOTTOM_RIGHT, $object->getAlignment());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->getColor());
        $this->assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
        $this->assertEquals(50, $object->getAlpha());
    }

    public function testSetGetAlignment ()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setAlignment());
        $this->assertEquals(0, $object->getAlignment());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setAlignment(Shadow::SHADOW_CENTER));
        $this->assertEquals(Shadow::SHADOW_CENTER, $object->getAlignment());
    }

    public function testSetGetAlpha ()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setAlpha());
        $this->assertEquals(0, $object->getAlpha());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setAlpha($value));
        $this->assertEquals($value, $object->getAlpha());
    }

    public function testSetGetBlurRadius ()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setBlurRadius());
        $this->assertEquals(6, $object->getBlurRadius());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setBlurRadius($value));
        $this->assertEquals($value, $object->getBlurRadius());
    }

    public function testSetGetColor ()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setColor());
        $this->assertNull($object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setColor(new Color(Color::COLOR_BLUE)));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->getColor());
        $this->assertEquals(Color::COLOR_BLUE, $object->getColor()->getARGB());
    }

    public function testSetGetDirection ()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setDirection());
        $this->assertEquals(0, $object->getDirection());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setDirection($value));
        $this->assertEquals($value, $object->getDirection());
    }

    public function testSetGetDistance ()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setDistance());
        $this->assertEquals(2, $object->getDistance());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setDistance($value));
        $this->assertEquals($value, $object->getDistance());
    }

    public function testSetGetHashIndex ()
    {
        $object = new Shadow();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
    
    public function testSetIsVisible ()
    {
        $object = new Shadow();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setVisible());
        $this->assertFalse($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setVisible(false));
        $this->assertFalse($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->setVisible(true));
        $this->assertTrue($object->isVisible());
    }
}
