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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;

/**
 * Test class for IOFactory
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\IOFactory
 */
class IOFactoryTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create writer
     */
    public function testCreateWriter()
    {
        $class = 'PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007';

        $this->assertInstanceOf($class, IOFactory::createWriter(new PhpPresentation()));
    }

    /**
     * Test create reader
     */
    public function testCreateReader()
    {
        $class = 'PhpOffice\\PhpPresentation\\Reader\\ReaderInterface';

        $this->assertInstanceOf($class, IOFactory::createReader('Serialized'));
    }

    /**
     * Test load class exception
     *
     * @expectedException \Exception
     * @expectedExceptionMessage is not a valid reader
     */
    public function testLoadClassException()
    {
        IOFactory::createReader();
    }

    public function testLoad()
    {
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', IOFactory::load(PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'files'.DIRECTORY_SEPARATOR.'serialized.phppt'));
    }

    /**
     * Test load class exception
     *
     * @expectedException \Exception
     * @expectedExceptionMessage Could not automatically determine \PhpOffice\PhpPresentation\Reader\ReaderInterface for file.
     */
    public function testLoadException()
    {
        IOFactory::load(PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png');
    }
}
