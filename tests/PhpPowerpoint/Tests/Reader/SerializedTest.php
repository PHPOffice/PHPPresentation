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

namespace PhpOffice\PhpPowerpoint\Tests\Reader;

use PhpOffice\PhpPowerpoint\Reader\Serialized;

/**
 * Test class for serialized reader
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Reader\Serialized
 */
class SerializedTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testCanRead()
    {
        $file = PHPPOWERPOINT_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        $object = new Serialized();

        $this->assertTrue($object->canRead($file));
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Could not open  for reading! File does not exist.
     */
    public function testLoadFileNotExists()
    {
        $object = new Serialized();
        $object->load('');
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Invalid file format for PhpOffice\PhpPowerpoint\Reader\Serialized: 
     */
    public function testLoadFileBadFormat()
    {
        $file = PHPPOWERPOINT_TESTS_BASE_DIR . '/resources/files/Sample_01_Simple.pptx';
        $object = new Serialized();
        $object->load($file);
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Could not open  for reading! File does not exist.
     */
    public function testFileSupportsNotExists()
    {
        $object = new Serialized();
        $object->fileSupportsUnserializePHPPowerPoint('');
    }
    
    public function testLoadSerializedFileNotExists()
    {
        $file = tempnam(sys_get_temp_dir(), 'PhpPowerpoint_Serialized');
        $oArchive = new \ZipArchive();
        $oArchive->open($file, \ZipArchive::CREATE);
        $oArchive->addFromString('PHPPowerPoint.xml', '');
        $oArchive->close();
        
        $object = new Serialized();
        $this->assertNull($object->load($file));
    }
}
