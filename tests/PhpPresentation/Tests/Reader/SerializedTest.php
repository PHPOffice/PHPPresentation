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
use PhpOffice\PhpPresentation\Reader\Serialized;
use PHPUnit\Framework\TestCase;
use ZipArchive;

/**
 * Test class for serialized reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Serialized
 */
class SerializedTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testCanRead(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        $object = new Serialized();

        $this->assertTrue($object->canRead($file));
    }

    public function testLoadFileNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new Serialized();
        $object->load('');
    }

    public function testLoadFileBadFormat(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_01_Simple.pptx';
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\Serialized',
            $file
        ));

        $object = new Serialized();
        $object->load($file);
    }

    public function testFileSupportsNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new Serialized();
        $object->fileSupportsUnserializePhpPresentation('');
    }

    public function testLoadSerializedFileNotExists(): void
    {
        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');

        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\Serialized (The file PhpPresentation.xml is malformed)',
            $file
        ));

        $oArchive = new ZipArchive();
        $oArchive->open($file, ZipArchive::CREATE);
        $oArchive->addFromString('PhpPresentation.xml', '');
        $oArchive->close();

        $object = new Serialized();
        $object->load($file);
    }
}
