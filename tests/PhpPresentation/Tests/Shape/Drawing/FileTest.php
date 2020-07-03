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

use PhpOffice\PhpPresentation\Shape\Drawing\File;

/**
 * Test class for Drawing element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Drawing
 */
class FileTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new File();
        $this->assertEmpty($object->getPath());
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage File  not found!
     */
    public function testPathBasic()
    {
        $object = new File();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\File', $object->setPath());
    }

    public function testPathWithoutVerifyFile()
    {
        $object = new File();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\File', $object->setPath('', false));
        $this->assertEmpty($object->getPath());
    }

    public function testPathWithRealFile()
    {
        $object = new File();

        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\File', $object->setPath($imagePath, false));
        $this->assertEquals($imagePath, $object->getPath());
        $this->assertEquals(0, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
    }
}
