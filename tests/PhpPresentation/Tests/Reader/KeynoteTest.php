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

namespace PhpOffice\PhpPresentation\Tests\Reader;

use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\Keynote;
use PhpOffice\PhpPresentation\Writer\Keynote as KeynoteWriter;
use PHPUnit\Framework\TestCase;

/**
 * Test class for the Keynote reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Keynote
 */
class KeynoteTest extends TestCase
{
    /**
     * Create a sample Keynote file on disk and return its path.
     */
    private function createSampleFile(): string
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('Sample text');

        $sFilename = tempnam(sys_get_temp_dir(), 'kny') . '.key';
        $oWriter = new KeynoteWriter($oPhpPresentation);
        $oWriter->save($sFilename);

        return $sFilename;
    }

    public function testCanReadValidFile(): void
    {
        $sFilename = $this->createSampleFile();

        $oReader = new Keynote();
        self::assertTrue($oReader->canRead($sFilename));

        unlink($sFilename);
    }

    public function testCanReadRejectsOtherFormats(): void
    {
        $oReader = new Keynote();
        self::assertFalse($oReader->canRead(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx'));
    }

    public function testCanReadRejectsNonZipFile(): void
    {
        $oReader = new Keynote();
        self::assertFalse($oReader->canRead(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt'));
    }

    public function testCanReadRejectsDirectory(): void
    {
        $oReader = new Keynote();
        self::assertFalse($oReader->canRead(PHPPRESENTATION_TESTS_BASE_DIR));
    }

    public function testLoadFileNotExistsThrows(): void
    {
        $this->expectException(FileNotFoundException::class);

        $oReader = new Keynote();
        $oReader->load(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/does_not_exist.key');
    }

    public function testLoadBadFormatThrows(): void
    {
        $this->expectException(InvalidFileFormatException::class);

        $oReader = new Keynote();
        $oReader->load(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx');
    }

    public function testLoadReadsContent(): void
    {
        $sFilename = $this->createSampleFile();

        $oReader = new Keynote();
        $oPhpPresentation = $oReader->load($sFilename);
        unlink($sFilename);

        self::assertCount(1, $oPhpPresentation->getAllSlides());
        $aShapes = $oPhpPresentation->getSlide(0)->getShapeCollection();
        self::assertCount(1, $aShapes);
    }
}
