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

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Writer\Serialized;
use PhpOffice\PhpPresentation\PhpPresentation;
use PHPUnit\Framework\TestCase;

/**
 * Test class for serialized reader
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Reader\Serialized
 */
class SerializedTest extends TestCase
{
    public function testConstruct()
    {
        $object = new Serialized(new PhpPresentation());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage No PhpPresentation assigned.
     */
    public function testEmptyConstruct()
    {
        $object = new Serialized();
        $object->getPhpPresentation();
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Filename is empty.
     */
    public function testSaveEmpty()
    {
        $object = new Serialized(new PhpPresentation());
        $object->save('');
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage No PhpPresentation assigned.
     */
    public function testSaveNoObject()
    {
        $object = new Serialized();
        $object->save('file.phpppt');
    }

    public function testSave()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png');
        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');
        $object = new Serialized($oPhpPresentation);
        $object->save($file);
        
        $this->assertFileExists($file);
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Could not open
     */
    public function testSaveNotExistingDir()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png');
        $object = new Serialized($oPhpPresentation);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');

        $this->assertFileExists($file, $object->save($file.DIRECTORY_SEPARATOR.'test'.DIRECTORY_SEPARATOR.'test'));
    }

    public function testSaveOverwriting()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png');
        
        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');
        file_put_contents($file, rand(1, 100));

        $object = new Serialized($oPhpPresentation);
        $object->save($file);
            
        $this->assertFileExists($file);
    }
}
