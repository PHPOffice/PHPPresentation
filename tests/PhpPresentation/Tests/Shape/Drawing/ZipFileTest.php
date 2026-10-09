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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Shape\Drawing;

use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Shape\Drawing\ZipFile;
use PHPUnit\Framework\TestCase;
use ZipArchive;

/**
 * Test class for Drawing element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Drawing
 */
class ZipFileTest extends TestCase
{
    /**
     * @var string
     */
    protected $fileOk;

    /**
     * @var string
     */
    protected $fileKoZip;

    /**
     * @var string
     */
    protected $fileKoFile;

    protected function setUp(): void
    {
        parent::setUp();

        $this->fileOk = 'zip://' . PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'files' . DIRECTORY_SEPARATOR . 'Sample_01_Simple.pptx#ppt/media/phppowerpoint_logo1.gif';
        $this->fileKoZip = 'zip://' . PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'fileNotExist.pptx#ppt/media/phppowerpoint_logo1.gif';
        $this->fileKoFile = 'zip://' . PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'files' . DIRECTORY_SEPARATOR . 'Sample_01_Simple.pptx#ppt/media/filenotexists.gif';
    }

    public function testContentsException(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage(sprintf(
            'The file "%s" doesn\'t exist',
            str_replace(['zip://', '#ppt/media/phppowerpoint_logo1.gif'], '', $this->fileKoZip)
        ));

        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileKoZip);
        $oDrawing->getContents();
    }

    public function testExtension(): void
    {
        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileOk);
        self::assertEquals('gif', $oDrawing->getExtension());
    }

    public function testMimeType(): void
    {
        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileOk);
        self::assertEquals('image/gif', $oDrawing->getMimeType());
    }

    public function testMetafile(): void
    {
        $path = (string) tempnam(sys_get_temp_dir(), 'PhpPresentation');
        $zip = new ZipArchive();
        $zip->open($path, ZipArchive::OVERWRITE);
        $zip->addFile(dirname(__DIR__, 4) . '/resources/images/inkscape_shapes_emfplus.emf', 'media/image1.emf');
        $zip->close();

        $oDrawing = new ZipFile();
        $oDrawing->setPath('zip://' . $path . '#media/image1.emf');
        self::assertEquals('emf', $oDrawing->getExtension());
        // `getimagesize()` knows no metafile : phpoffice/wmf says what it is
        self::assertEquals('image/x-emf', $oDrawing->getMimeType());

        unlink($path);
    }
}
