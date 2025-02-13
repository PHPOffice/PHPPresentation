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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\PresentationProperties;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;

/**
 * Test class for DocumentProperties.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PresentationProperties
 */
class PresentationPropertiesTest extends TestCase
{
    public function testCommentVisible(): void
    {
        $object = new PresentationProperties();
        self::assertFalse($object->isCommentVisible());
        self::assertInstanceOf(PresentationProperties::class, $object->setCommentVisible(true));
        self::assertTrue($object->isCommentVisible());
        self::assertInstanceOf(PresentationProperties::class, $object->setCommentVisible(false));
        self::assertFalse($object->isCommentVisible());
        self::assertInstanceOf(PresentationProperties::class, $object->setCommentVisible());
        self::assertFalse($object->isCommentVisible());
    }

    public function testLoopUntilEsc(): void
    {
        $object = new PresentationProperties();
        self::assertFalse($object->isLoopContinuouslyUntilEsc());
        self::assertInstanceOf(PresentationProperties::class, $object->setLoopContinuouslyUntilEsc(true));
        self::assertTrue($object->isLoopContinuouslyUntilEsc());
        self::assertInstanceOf(PresentationProperties::class, $object->setLoopContinuouslyUntilEsc(false));
        self::assertFalse($object->isLoopContinuouslyUntilEsc());
        self::assertInstanceOf(PresentationProperties::class, $object->setLoopContinuouslyUntilEsc());
        self::assertFalse($object->isLoopContinuouslyUntilEsc());
    }

    public function testLastView(): void
    {
        $object = new PresentationProperties();
        self::assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
        self::assertInstanceOf(PresentationProperties::class, $object->setLastView('AAAA'));
        self::assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
        self::assertInstanceOf(PresentationProperties::class, $object->setLastView(PresentationProperties::VIEW_OUTLINE));
        self::assertEquals(PresentationProperties::VIEW_OUTLINE, $object->getLastView());
        self::assertInstanceOf(PresentationProperties::class, $object->setLastView());
        self::assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
    }

    public function testMarkAsFinal(): void
    {
        $object = new PresentationProperties();
        self::assertFalse($object->isMarkedAsFinal());
        self::assertInstanceOf(PresentationProperties::class, $object->markAsFinal(true));
        self::assertTrue($object->isMarkedAsFinal());
        self::assertInstanceOf(PresentationProperties::class, $object->markAsFinal(false));
        self::assertFalse($object->isMarkedAsFinal());
        self::assertInstanceOf(PresentationProperties::class, $object->markAsFinal());
        self::assertTrue($object->isMarkedAsFinal());
    }

    /**
     * @dataProvider dataProviderSlideshowType
     */
    #[DataProvider('dataProviderSlideshowType')]
    public function testSlideshowType(?string $value, string $expected): void
    {
        $object = new PresentationProperties();
        // Default
        self::assertEquals(PresentationProperties::SLIDESHOW_TYPE_PRESENT, $object->getSlideshowType());
        // Set value
        if (null === $value) {
            self::assertInstanceOf(PresentationProperties::class, $object->setSlideshowType());
        } else {
            self::assertInstanceOf(PresentationProperties::class, $object->setSlideshowType($value));
        }
        // Check value
        self::assertEquals($expected, $object->getSlideshowType());
    }

    /**
     * @return array<array<null|string>>
     */
    public static function dataProviderSlideshowType(): array
    {
        return [
            [
                null,
                PresentationProperties::SLIDESHOW_TYPE_PRESENT,
            ],
            [
                PresentationProperties::SLIDESHOW_TYPE_PRESENT,
                PresentationProperties::SLIDESHOW_TYPE_PRESENT,
            ],
            [
                PresentationProperties::SLIDESHOW_TYPE_KIOSK,
                PresentationProperties::SLIDESHOW_TYPE_KIOSK,
            ],
            [
                PresentationProperties::SLIDESHOW_TYPE_BROWSE,
                PresentationProperties::SLIDESHOW_TYPE_BROWSE,
            ],
            [
                'unauthorizedValue',
                PresentationProperties::SLIDESHOW_TYPE_PRESENT,
            ],
        ];
    }

    public function testThumbnail(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        $object = new PresentationProperties();
        self::assertNull($object->getThumbnailPath());
        self::assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath('AAAA', PresentationProperties::THUMBNAIL_FILE));
        self::assertNull($object->getThumbnailPath());
        self::assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath());
        self::assertNull($object->getThumbnailPath());
        self::assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath($imagePath, PresentationProperties::THUMBNAIL_FILE));
        self::assertEquals($imagePath, $object->getThumbnailPath());
        self::assertIsString($object->getThumbnail());
        self::assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath());
        self::assertEquals($imagePath, $object->getThumbnailPath());
        self::assertIsString($object->getThumbnail());
    }

    public function testThumbnailFileNotExisting(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'NotExistingFile.png';

        $object = new PresentationProperties();
        self::assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath($imagePath, PresentationProperties::THUMBNAIL_FILE));
        self::assertNull($object->getThumbnailPath());
        self::assertNull($object->getThumbnail());
    }

    public function testThumbnailFileData(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        $object = new PresentationProperties();
        self::assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath($imagePath, PresentationProperties::THUMBNAIL_DATA, file_get_contents($imagePath)));
        self::assertEquals('', $object->getThumbnailPath());
        self::assertIsString($object->getThumbnail());
    }

    public function testZoom(): void
    {
        $object = new PresentationProperties();
        self::assertEquals(1, $object->getZoom());
        self::assertInstanceOf(PresentationProperties::class, $object->setZoom(2.3));
        self::assertEquals(2.3, $object->getZoom());
        self::assertInstanceOf(PresentationProperties::class, $object->setZoom());
        self::assertEquals(1, $object->getZoom());
    }
}
