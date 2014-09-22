<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Writer;

use PhpOffice\PhpPowerpoint\Writer\Serialized;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;

/**
 * Test class for serialized reader
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Reader\Serialized
 */
class SerializedTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Serialized(new PhpPowerpoint());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\PhpPowerpoint', $object->getPHPPowerPoint());
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage No PHPPowerPoint assigned.
     */
    public function testEmptyConstruct()
    {
        $object = new Serialized();
        $object->getPHPPowerPoint();
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Filename is empty.
     */
    public function testSaveEmpty()
    {
        $object = new Serialized(new PhpPowerpoint());
        $object->save('');
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage PHPPowerPoint object unassigned.
     */
    public function testSaveNoObject()
    {
        $object = new Serialized();
        $object->save('file.phpppt');
    }
    
    public function testSave()
    {
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPOWERPOINT_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PHPPowerPointLogo.png');
        $object = new Serialized($oPHPPowerPoint);
        
        $file = tempnam(sys_get_temp_dir(), 'PhpPowerpoint_Serialized');
        
        $this->assertFileExists($file, $object->save($file));
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Could not open 
     */
    public function testSaveNotExistingDir()
    {
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPOWERPOINT_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PHPPowerPointLogo.png');
        $object = new Serialized($oPHPPowerPoint);
        
        $file = tempnam(sys_get_temp_dir(), 'PhpPowerpoint_Serialized');
        
        $this->assertFileExists($file, $object->save($file.DIRECTORY_SEPARATOR.'test'.DIRECTORY_SEPARATOR.'test'));
    }
    
    public function testSaveOverwriting()
    {
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPOWERPOINT_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PHPPowerPointLogo.png');
        $object = new Serialized($oPHPPowerPoint);
        
        $file = tempnam(sys_get_temp_dir(), 'PhpPowerpoint_Serialized');
        file_put_contents($file, rand(1, 100));
        
        $this->assertFileExists($file, $object->save($file));
    }
}
