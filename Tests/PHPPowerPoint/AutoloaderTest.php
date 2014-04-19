<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

/**
 * Test class for PHPPowerPoint_Autoloader
 */
class AutoloaderTest extends PHPUnit_Framework_TestCase
{
    /**
     * Test register
     */
    public function testRegister()
    {
        $this->assertTrue(PHPPowerPoint_Autoloader::Register());
    }

    /**
     * Test load
     */
    public function testAutoloaderNonPHPPowerPointClass()
    {
        $className = 'InvalidClass';

        $result = PHPPowerPoint_Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    /**
     * Test load
     */
    public function testAutoloaderInvalidPHPPowerPointClass()
    {
        $className = 'PHPPowerPoint_Invalid_Class';

        $result = PHPPowerPoint_Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    /**
     * Test load
     */
    public function testAutoloadValidPHPPowerPointClass()
    {
        $className = 'PHPPowerPoint_IOFactory';

        $result = PHPPowerPoint_Autoloader::Load($className);
        //    Check that class has been loaded
        $this->assertTrue(class_exists($className));
    }

    /**
     * Test load
     */
    public function testAutoloadInstantiateSuccess()
    {
        $result = new PHPPowerPoint();
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, 'PHPPowerPoint'));
    }
}
