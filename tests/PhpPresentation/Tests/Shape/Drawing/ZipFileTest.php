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

namespace PhpOffice\PhpPresentation\Tests\Shape\Drawing;

use PhpOffice\PhpPresentation\Shape\Drawing\ZipFile;

/**
 * Test class for Drawing element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Drawing
 */
class ZipFileTest extends \PHPUnit_Framework_TestCase
{
    protected $fileOk;
    protected $fileKoZip;
    protected $fileKoFile;

    public function setUp()
    {
        parent::setUp();

        DrawingTest::$getimagesizefromstringExists = true;

        $this->fileOk = 'zip://'.PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'files'.DIRECTORY_SEPARATOR.'Sample_01_Simple.pptx#ppt/media/phppowerpoint_logo1.gif';
        $this->fileKoZip = 'zip://'.PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'fileNotExist.pptx#ppt/media/phppowerpoint_logo1.gif';
        $this->fileKoFile = 'zip://'.PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'files'.DIRECTORY_SEPARATOR.'Sample_01_Simple.pptx#ppt/media/filenotexists.gif';
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage fileNotExist.pptx does not exist
     */
    public function testContentsException()
    {
        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileKoZip);
        $oDrawing->getContents();
    }

    public function testExtension()
    {
        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileOk);
        $this->assertEquals('gif', $oDrawing->getExtension());
    }

    /**
     * @requires PHP 5.4
     */
    public function testMimeType()
    {
        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileOk);
        $this->assertEquals('image/gif', $oDrawing->getMimeType());
    }

    /**
     * @requires PHP 5.4
     */
    public function testMimeTypeFunctionNotExists()
    {
        DrawingTest::$getimagesizefromstringExists = false;
        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileOk);
        $this->assertEquals('image/gif', $oDrawing->getMimeType());
    }
}
