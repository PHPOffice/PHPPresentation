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

namespace PhpOffice\PhpPresentation\Tests\Shared;

use PhpOffice\PhpPresentation\Shared\Metafile;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpOffice\PhpPresentation\Shared\Metafile.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shared\Metafile
 */
class MetafileTest extends TestCase
{
    public function testIsSupported(): void
    {
        self::assertTrue(Metafile::isSupported());
    }

    /**
     * @dataProvider dataProviderMetafile
     */
    #[DataProvider('dataProviderMetafile')]
    public function testMetafile(string $filename, string $mimeType, string $extension, int $width, int $height): void
    {
        $contents = (string) file_get_contents(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/' . $filename);

        self::assertEquals($mimeType, Metafile::getMimeType($contents));
        self::assertTrue(Metafile::isMimeType($mimeType));
        self::assertEquals($extension, Metafile::getExtension($mimeType));
        self::assertEquals([$width, $height, $mimeType], Metafile::getImageSize($contents));

        $png = Metafile::convertToPng($contents);
        self::assertIsString($png);
        $size = getimagesizefromstring($png);
        self::assertIsArray($size);
        self::assertEquals('image/png', $size['mime']);
        self::assertEquals($width, $size[0]);
        self::assertEquals($height, $size[1]);
    }

    /**
     * @return array<array{string, string, string, int, int}>
     */
    public static function dataProviderMetafile(): array
    {
        return [
            'WMF' => ['fish.wmf', Metafile::MIME_WMF, 'wmf', 217, 159],
            'EMF' => ['inkscape_shapes.emf', Metafile::MIME_EMF, 'emf', 200, 151],
            // an EMF+ file is an EMF file, and is stored as one
            'EMF+' => ['inkscape_shapes_emfplus.emf', Metafile::MIME_EMF, 'emf', 200, 151],
        ];
    }

    public function testNotAMetafile(): void
    {
        $contents = (string) file_get_contents(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        self::assertNull(Metafile::getMimeType($contents));
        self::assertNull(Metafile::getImageSize($contents));
        self::assertNull(Metafile::convertToPng($contents));
        self::assertNull(Metafile::getMimeType(''));
    }

    public function testMimeTypes(): void
    {
        // the types without their `x-` are met too
        self::assertEquals('wmf', Metafile::getExtension('image/wmf'));
        self::assertEquals('emf', Metafile::getExtension('image/emf'));
        self::assertEquals('emf', Metafile::getExtension('IMAGE/X-EMF'));
        self::assertNull(Metafile::getExtension('image/png'));
        self::assertFalse(Metafile::isMimeType('image/png'));
    }
}
