<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Style\Color;

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\PhpPowerpoint
 */
class BorderTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct ()
    {
        $object = new Border();
        $this->assertEquals(1, $object->getLineWidth());
        $this->assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        $this->assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->getColor());
        $this->assertEquals('FF000000', $object->getColor()->getARGB());
    }

    /**
     * Test get/set color
     */
    public function testSetGetColor ()
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setColor());
        $this->assertNull($object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setColor(new Color(COLOR::COLOR_BLUE)));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Color', $object->getColor());
        $this->assertEquals('FF0000FF', $object->getColor()->getARGB());
    }

    /**
     * Test get/set dash style
     */
    public function testSetGetDashStyle ()
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setDashStyle());
        $this->assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setDashStyle(''));
        $this->assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setDashStyle(BORDER::DASH_DASH));
        $this->assertEquals(Border::DASH_DASH, $object->getDashStyle());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex ()
    {
        $object = new Border();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set line style
     */
    public function testSetGetLineStyle ()
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setLineStyle());
        $this->assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setLineStyle(''));
        $this->assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setLineStyle(BORDER::LINE_DOUBLE));
        $this->assertEquals(Border::LINE_DOUBLE, $object->getLineStyle());
    }

    /**
     * Test get/set line width
     */
    public function testSetGetLineWidth ()
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setLineWidth());
        $this->assertEquals(1, $object->getLineWidth());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->setLineWidth($value));
        $this->assertEquals($value, $object->getLineWidth());
    }
}
