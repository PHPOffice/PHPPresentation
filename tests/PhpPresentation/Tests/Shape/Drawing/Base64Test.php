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

use PhpOffice\PhpPresentation\Exception\UnauthorizedMimetypeException;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Drawing element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Drawing
 */
class Base64Test extends TestCase
{
    /** @var string */
    protected $imageDataPNG = '';

    /** @var string */
    protected $imageDataSVG = '';

    protected function setUp(): void
    {
        parent::setUp();

        $this->imageDataPNG = file_get_contents(dirname(__DIR__, 4) . '/resources/images/base64_png.txt');

        $this->imageDataSVG = file_get_contents(dirname(__DIR__, 4) . '/resources/images/base64_svg.txt');
    }

    public function testData(): void
    {
        $oDrawing = new Base64();

        self::assertEmpty($oDrawing->getData());
        $oDrawing->setData($this->imageDataPNG);
        self::assertNotEmpty($oDrawing->getData());
    }

    public function testExtension(): void
    {
        $oDrawing = new Base64();
        $oDrawing->setData($this->imageDataPNG);
        self::assertEquals('jpg', $oDrawing->getExtension());
    }

    public function testExtensionException(): void
    {
        $this->expectException(UnauthorizedMimetypeException::class);
        $this->expectExceptionMessage('The mime type fake/fake is not found in autorized values (jpg, png, gif, svg, wmf, emf)');

        $imgData = str_replace('image/jpeg', 'fake/fake', $this->imageDataPNG);

        $oDrawing = new Base64();
        $oDrawing->setData($imgData);
        $oDrawing->getExtension();
    }

    public function testMimeType(): void
    {
        $oDrawing = new Base64();
        $oDrawing->setData($this->imageDataPNG);
        self::assertEquals('image/jpeg', $oDrawing->getMimeType());
    }

    public function testMimeTypeSVG(): void
    {
        $oDrawing = new Base64();
        $oDrawing->setData($this->imageDataSVG);
        self::assertEquals('image/svg+xml', $oDrawing->getMimeType());
    }

    public function testMetafile(): void
    {
        $contents = (string) file_get_contents(dirname(__DIR__, 4) . '/resources/images/fish.wmf');

        $oDrawing = new Base64();
        $oDrawing->setData('data:image/x-wmf;base64,' . base64_encode($contents));
        self::assertEquals('wmf', $oDrawing->getExtension());
        self::assertEquals('image/x-wmf', $oDrawing->getMimeType());
        self::assertEquals($contents, $oDrawing->getContents());

        // the types without their `x-` are met too
        $oDrawing->setData('data:image/emf;base64,' . base64_encode($contents));
        self::assertEquals('emf', $oDrawing->getExtension());
    }

    public function testMetafileWithoutMimeType(): void
    {
        $contents = (string) file_get_contents(dirname(__DIR__, 4) . '/resources/images/inkscape_shapes.emf');

        $oDrawing = new Base64();
        $oDrawing->setData('data:;base64,' . base64_encode($contents));
        self::assertEquals('image/x-emf', $oDrawing->getMimeType());
    }

    public function testMimeTypeFunctionNotExists(): void
    {
        $oDrawing = new Base64();
        $oDrawing->setData($this->imageDataPNG);
        self::assertEquals('image/jpeg', $oDrawing->getMimeType());
    }
}
