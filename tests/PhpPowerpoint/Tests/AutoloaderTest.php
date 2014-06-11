<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\Autoloader;

/**
 * Test class for Autoloader
 */
class AutoloaderTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Register
     */
    public function testRegister()
    {
        Autoloader::register();
        $this->assertContains(
            array('PhpOffice\\PhpPowerpoint\\Autoloader', 'autoload'),
            spl_autoload_functions()
        );
    }
    
    /**
     * Autoload
     */
    public function testAutoload()
    {
        $declared = get_declared_classes();
        $declaredCount = count($declared);
        Autoloader::autoload('Foo');
        $this->assertEquals(
            $declaredCount,
            count(get_declared_classes()),
            'PhpOffice\\PhpPowerpoint\\Autoloader::autoload() is trying to load ' .
            'classes outside of the PhpOffice\\PhpPowerpoint namespace'
        );
        // TODO change this class to the main PhpPowerpoint class when it is namespaced
        Autoloader::autoload('PhpOffice\\PhpPowerpoint\\Exception\\InvalidStyleException');
        $this->assertTrue(
            in_array('PhpOffice\\PhpPowerpoint\\Exception\\InvalidStyleException', get_declared_classes()),
            'PhpOffice\\PhpPowerpoint\\Autoloader::autoload() failed to autoload the ' .
            'PhpOffice\\PhpPowerpoint\\Exception\\InvalidStyleException class'
        );
    }
}
