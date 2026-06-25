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

namespace PhpOffice\PhpPresentation\Tests\Writer;

use DOMDocument;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Writer\Keynote;
use PHPUnit\Framework\TestCase;
use ZipArchive;

/**
 * Test class for the Keynote writer.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\Keynote
 */
class KeynoteTest extends TestCase
{
    public function testConstruct(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oWriter = new Keynote($oPhpPresentation);
        self::assertSame($oPhpPresentation, $oWriter->getPhpPresentation());
    }

    public function testSaveCreatesApxlEntry(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('Hello Keynote');

        $sFilename = tempnam(sys_get_temp_dir(), 'kny') . '.key';
        $oWriter = new Keynote($oPhpPresentation);
        $oWriter->save($sFilename);

        self::assertFileExists($sFilename);

        $oZip = new ZipArchive();
        self::assertTrue(true === $oZip->open($sFilename));
        $sContent = $oZip->getFromName('index.apxl');
        $oZip->close();

        self::assertNotFalse($sContent);

        $oDom = new DOMDocument();
        self::assertTrue($oDom->loadXML((string) $sContent));
        self::assertStringContainsString('Hello Keynote', (string) $sContent);
        self::assertStringContainsString('key:presentation', (string) $sContent);

        unlink($sFilename);
    }

    public function testSaveWithEmptyFilenameThrows(): void
    {
        $this->expectException(InvalidParameterException::class);

        $oWriter = new Keynote(new PhpPresentation());
        $oWriter->save('');
    }
}
