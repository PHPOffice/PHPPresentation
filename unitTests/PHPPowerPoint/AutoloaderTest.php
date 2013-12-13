<?php


class AutoloaderTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        if (!defined('PHPPOWERPOINT_ROOT'))
        {
            define('PHPPOWERPOINT_ROOT', APPLICATION_PATH . '/');
        }
        require_once(PHPPOWERPOINT_ROOT . 'PHPPowerPoint/Autoloader.php');
    }

    public function testAutoloaderNonPHPPowerPointClass()
    {
        $className = 'InvalidClass';

        $result = PHPPowerPoint_Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    public function testAutoloaderInvalidPHPPowerPointClass()
    {
        $className = 'PHPPowerPoint_Invalid_Class';

        $result = PHPPowerPoint_Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    public function testAutoloadValidPHPPowerPointClass()
    {
        $className = 'PHPPowerPoint_IOFactory';

        $result = PHPPowerPoint_Autoloader::Load($className);
        //    Check that class has been loaded
        $this->assertTrue(class_exists($className));
    }

    public function testAutoloadInstantiateSuccess()
    {
        $result = new PHPPowerPoint(1,2,3);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result,'PHPPowerPoint'));
    }

}