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

namespace PhpOffice\PhpPresentation\Tests\Reader;

use PhpOffice\PhpPresentation\Reader\Serialized;

/**
 * Test class for serialized reader
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Reader\Serialized
 */
class SerializedTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testCanRead()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
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
     * @expectedExceptionMessage Invalid file format for PhpOffice\PhpPresentation\Reader\Serialized:
     */
    public function testLoadFileBadFormat()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_01_Simple.pptx';
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
        $object->fileSupportsUnserializePhpPresentation('');
    }
    
    public function testLoadSerializedFileNotExists()
    {
        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');
        $oArchive = new \ZipArchive();
        $oArchive->open($file, \ZipArchive::CREATE);
        $oArchive->addFromString('PhpPresentation.xml', '');
        $oArchive->close();
        
        $object = new Serialized();
        $this->assertNull($object->load($file));
    }
}
