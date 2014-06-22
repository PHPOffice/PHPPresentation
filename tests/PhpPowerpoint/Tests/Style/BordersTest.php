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
use PhpOffice\PhpPowerpoint\Style\Borders;

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\PhpPowerpoint
 */
class BordersTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct ()
    {
        $object = new Borders();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getBottom());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getLeft());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getRight());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getTop());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getDiagonalDown());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getDiagonalUp());
        $this->assertEquals(Border::LINE_NONE, $object->getDiagonalDown()->getLineStyle());
        $this->assertEquals(Border::LINE_NONE, $object->getDiagonalUp()->getLineStyle());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex ()
    {
        $object = new Borders();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set hash code
     */
    public function testGetHashCode ()
    {
        $object = new Borders();
        $this->assertEquals(
            md5(
                $object->getLeft()->getHashCode() .
                $object->getRight()->getHashCode() .
                $object->getTop()->getHashCode() .
                $object->getBottom()->getHashCode() .
                $object->getDiagonalUp()->getHashCode() .
                $object->getDiagonalDown()->getHashCode() .
                get_class($object)
            ),
            $object->getHashCode()
        );
    }
}
