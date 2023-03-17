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

namespace PhpOffice\PhpPresentation\Tests\Shape\Drawing;

use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Shape\Drawing\ZipFile;
use PHPUnit\Framework\TestCase;

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
        $this->assertEquals('gif', $oDrawing->getExtension());
    }

    public function testMimeType(): void
    {
        $oDrawing = new ZipFile();
        $oDrawing->setPath($this->fileOk);
        $this->assertEquals('image/gif', $oDrawing->getMimeType());
    }
}
