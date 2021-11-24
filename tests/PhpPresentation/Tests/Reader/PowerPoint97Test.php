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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Reader;

use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\Reader\PowerPoint97;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PowerPoint97 reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\PowerPoint97
 */
class PowerPoint97Test extends TestCase
{
    /**
     * Test can read.
     */
    public function testCanRead(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        $object = new PowerPoint97();

        $this->assertTrue($object->canRead($file));
    }

    /**
     * Test cant read.
     */
    public function testCantRead(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        $object = new PowerPoint97();

        $this->assertFalse($object->canRead($file));
    }

    public function testLoadFileNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new PowerPoint97();
        $object->load('');
    }

    public function testLoadFileBadFormat(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_01_Simple.pptx';
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\PowerPoint97',
            $file
        ));

        $object = new PowerPoint97();
        $object->load($file);
    }

    public function testFileSupportsNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new PowerPoint97();
        $object->fileSupportsUnserializePhpPresentation('');
    }

    public function testLoadFile01(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        $object = new PowerPoint97();
        $oPhpPresentation = $object->load($file);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        $this->assertEquals(1, $oPhpPresentation->getSlideCount());

        $oSlide = $oPhpPresentation->getSlide(0);
        $this->assertCount(2, $oSlide->getShapeCollection());
    }

    public function testLoadFile02(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_02.ppt';
        $object = new PowerPoint97();
        $oPhpPresentation = $object->load($file);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        $this->assertEquals(4, $oPhpPresentation->getSlideCount());

        $oSlide = $oPhpPresentation->getSlide(0);
        $this->assertCount(2, $oSlide->getShapeCollection());

        $oSlide = $oPhpPresentation->getSlide(1);
        $this->assertCount(3, $oSlide->getShapeCollection());

        $oSlide = $oPhpPresentation->getSlide(2);
        $this->assertCount(3, $oSlide->getShapeCollection());

        $oSlide = $oPhpPresentation->getSlide(3);
        $this->assertCount(3, $oSlide->getShapeCollection());
    }

    public function testLoadFile03(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_03.ppt';
        $object = new PowerPoint97();
        $oPhpPresentation = $object->load($file);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        $this->assertEquals(1, $oPhpPresentation->getSlideCount());

        $oSlide = $oPhpPresentation->getSlide(0);
        $this->assertCount(1, $oSlide->getShapeCollection());
    }

    public function testLoadFile04(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_04.ppt';
        $object = new PowerPoint97();
        $oPhpPresentation = $object->load($file);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        $this->assertEquals(1, $oPhpPresentation->getSlideCount());

        $oSlide = $oPhpPresentation->getSlide(0);
        $this->assertCount(4, $oSlide->getShapeCollection());
    }
}
