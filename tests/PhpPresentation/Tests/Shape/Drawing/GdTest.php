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
 * @license     http://www.gnu.org/licenses/lgpl.html LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Shape\Drawing;

use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Drawing element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Drawing
 */
class GdTest extends TestCase
{
    public function testContentsAreTheBytesThePictureWasReadFrom(): void
    {
        $oDrawing = $this->drawing();

        $bytes = 'the bytes the file holds';
        $oDrawing->setContents($bytes);

        self::assertSame($bytes, $oDrawing->getContents());
    }

    public function testContentsAreRenderedWhenThereAreNoBytes(): void
    {
        $oDrawing = $this->drawing();

        self::assertNotEmpty($oDrawing->getContents());
    }

    public function testSettingAResourceDropsTheBytes(): void
    {
        $oDrawing = $this->drawing();
        $oDrawing->setContents('the bytes the file holds');

        $other = imagecreatetruecolor(8, 8);
        $oDrawing->setImageResource($other);

        self::assertNotSame('the bytes the file holds', $oDrawing->getContents());
    }

    public function testIndexedFilenameIsTheFileThePictureWasReadFrom(): void
    {
        $oDrawing = new Gd();
        $oDrawing->setPath('zip:///path/to/sample.pptx#ppt/media/image1.png');

        // The index is a counter of its own, so the name is what says it and not the
        // whole of it.
        self::assertStringStartsWith('image1', $oDrawing->getIndexedFilename());
        self::assertStringEndsWith('.png', $oDrawing->getIndexedFilename());
    }

    public function testIndexedFilenameIsMadeUpWhenThereIsNoPath(): void
    {
        $oDrawing = $this->drawing();

        self::assertStringEndsWith('.png', $oDrawing->getIndexedFilename());
        self::assertStringNotContainsString('image1', $oDrawing->getIndexedFilename());
    }

    public function testTwoPicturesWithTheSameFileNameAreNamedApart(): void
    {
        $first = new Gd();
        $first->setPath('zip:///path/to/sample.pptx#media/1/logo.png');

        $second = new Gd();
        $second->setPath('zip:///path/to/sample.pptx#media/2/logo.png');

        self::assertNotSame($first->getIndexedFilename(), $second->getIndexedFilename());
    }

    private function drawing(): Gd
    {
        $oDrawing = new Gd();
        $image = imagecreatetruecolor(8, 8);
        imagefill($image, 0, 0, imagecolorallocate($image, 200, 30, 30));
        $oDrawing->setImageResource($image);

        return $oDrawing;
    }
}
