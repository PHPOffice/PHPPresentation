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

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\PhpPowerpoint
 */
class ColorTest    extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct ()
    {
        $object = new Color();
        $this->assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $object = new Color(COLOR::COLOR_BLUE);
        $this->assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }

    public function testSetGetARGB ()
    {
        $object = new Color();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->setARGB());
        $this->assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->setARGB(''));
        $this->assertEquals(Color::COLOR_BLACK, $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->setARGB(Color::COLOR_BLUE));
        $this->assertEquals(Color::COLOR_BLUE, $object->getARGB());
    }
    public function testSetGetRGB ()
    {
        $object = new Color();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->setRGB());
        $this->assertEquals('000000', $object->getRGB());
        $this->assertEquals('FF000000', $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->setRGB(''));
        $this->assertEquals('000000', $object->getRGB());
        $this->assertEquals('FF000000', $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->setRGB('555'));
        $this->assertEquals('555', $object->getRGB());
        $this->assertEquals('FF555', $object->getARGB());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->setRGB('6666'));
        $this->assertEquals('FF6666', $object->getRGB());
        $this->assertEquals('FF6666', $object->getARGB());
    }
    
    public function testSetGetHashIndex ()
    {
        $object = new Color();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
}
