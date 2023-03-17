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
use PhpOffice\PhpPresentation\Shape\Drawing\File;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Drawing element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Drawing
 */
class FileTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new File();
        $this->assertEmpty($object->getPath());
    }

    public function testPathBasic(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new File();
        $this->assertInstanceOf(File::class, $object->setPath());
    }

    public function testPathWithoutVerifyFile(): void
    {
        $object = new File();

        $this->assertInstanceOf(File::class, $object->setPath('', false));
        $this->assertEmpty($object->getPath());
    }

    public function testPathWithRealFile(): void
    {
        $object = new File();

        $imagePath = dirname(__DIR__, 4) . '/resources/images/PhpPresentationLogo.png';

        $this->assertInstanceOf(File::class, $object->setPath($imagePath, false));
        $this->assertEquals($imagePath, $object->getPath());
        $this->assertEquals(0, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
    }

    /**
     * @dataProvider dataProviderMimeType
     */
    public function testMimeType(string $pathFile, string $mimeType): void
    {
        $object = new File();
        $this->assertInstanceOf(File::class, $object->setPath($pathFile));
        $this->assertEquals($mimeType, $object->getMimeType());
    }

    /**
     * @return array<array<string>>
     */
    public function dataProviderMimeType(): array
    {
        return [
            [
                dirname(__DIR__, 4) . '/resources/images/PhpPresentationLogo.png',
                'image/png',
            ],
            [
                dirname(__DIR__, 4) . '/resources/images/tiger.svg',
                'image/svg+xml',
            ],
        ];
    }
}
