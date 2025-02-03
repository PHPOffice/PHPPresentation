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

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PHPUnit\Framework\TestCase;

class ImageTest extends TestCase
{
    public function testColor(): void
    {
        $object = new Image();

        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';
        $numSlide = (string) mt_rand(1, 100);

        self::assertNull($object->getPath());
        self::assertEmpty($object->getFilename());
        self::assertEmpty($object->getExtension());
        self::assertEquals('background_' . $numSlide . '.', $object->getIndexedFilename($numSlide));

        self::assertInstanceOf(Image::class, $object->setPath($imagePath));
        self::assertEquals($imagePath, $object->getPath());
        self::assertEquals('PhpPresentationLogo.png', $object->getFilename());
        self::assertEquals('png', $object->getExtension());
        self::assertEquals('background_' . $numSlide . '.png', $object->getIndexedFilename($numSlide));

        self::assertInstanceOf(Image::class, $object->setPath('', false));
        self::assertEquals('', $object->getPath());
        self::assertEmpty($object->getFilename());
        self::assertEmpty($object->getExtension());
        self::assertEquals('background_' . $numSlide . '.', $object->getIndexedFilename($numSlide));
    }

    public function testExtension(): void
    {
        $object = new Image();
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        self::assertEmpty($object->getExtension());
        self::assertInstanceOf(Image::class, $object->setPath($imagePath));
        self::assertEquals('png', $object->getExtension());
        self::assertInstanceOf(Image::class, $object->setExtension('jpg'));
        self::assertEquals('jpg', $object->getExtension());
    }

    public function testPathException(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "pathDoesntExist" doesn\'t exist');

        $object = new Image();
        $object->setPath('pathDoesntExist', true);
    }
}
