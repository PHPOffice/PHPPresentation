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

namespace PhpOffice\PhpPowerpoint\Tests\Shared;

use PhpOffice\PhpPowerpoint\Shared\File;

/**
 * Test class for IOFactory
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\IOFactory
 */
class FileTest extends \PHPUnit_Framework_TestCase
{
    /**
     */
    public function testFileExists()
    {
        $pathResources = PHPPOWERPOINT_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR;
        $this->assertTrue(File::fileExists($pathResources.'images'.DIRECTORY_SEPARATOR.'PHPPowerPointLogo.png'));
        $this->assertFalse(File::fileExists($pathResources.'images'.DIRECTORY_SEPARATOR.'PHPPowerPointLogo_404.png'));
        $this->assertTrue(File::fileExists('zip://'.$pathResources.'files'.DIRECTORY_SEPARATOR.'Sample_01_Simple.pptx#[Content_Types].xml'));
        $this->assertFalse(File::fileExists('zip://'.$pathResources.'files'.DIRECTORY_SEPARATOR.'Sample_01_Simple.pptx#404.xml'));
        $this->assertFalse(File::fileExists('zip://'.$pathResources.'files'.DIRECTORY_SEPARATOR.'404.pptx#404.xml'));
    }

    /**
     */
    public function testRealPath()
    {
        $pathFiles = PHPPOWERPOINT_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'files'.DIRECTORY_SEPARATOR;

        $this->assertEquals($pathFiles.'Sample_01_Simple.pptx', File::realpath($pathFiles.'Sample_01_Simple.pptx'));
        $this->assertEquals('zip://'.$pathFiles.'Sample_01_Simple.pptx#[Content_Types].xml', File::realpath('zip://'.$pathFiles.'Sample_01_Simple.pptx#[Content_Types].xml'));
        $this->assertEquals('zip://'.$pathFiles.'Sample_01_Simple.pptx#/[Content_Types].xml', File::realpath('zip://'.$pathFiles.'Sample_01_Simple.pptx#/rels/../[Content_Types].xml'));
    }
}
