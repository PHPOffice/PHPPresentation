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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\PresentationProperties;
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
        $this->assertFalse($object->isCommentVisible());
        $this->assertInstanceOf(PresentationProperties::class, $object->setCommentVisible(true));
        $this->assertTrue($object->isCommentVisible());
        $this->assertInstanceOf(PresentationProperties::class, $object->setCommentVisible(false));
        $this->assertFalse($object->isCommentVisible());
        $this->assertInstanceOf(PresentationProperties::class, $object->setCommentVisible());
        $this->assertFalse($object->isCommentVisible());
    }

    public function testLoopUntilEsc(): void
    {
        $object = new PresentationProperties();
        $this->assertFalse($object->isLoopContinuouslyUntilEsc());
        $this->assertInstanceOf(PresentationProperties::class, $object->setLoopContinuouslyUntilEsc(true));
        $this->assertTrue($object->isLoopContinuouslyUntilEsc());
        $this->assertInstanceOf(PresentationProperties::class, $object->setLoopContinuouslyUntilEsc(false));
        $this->assertFalse($object->isLoopContinuouslyUntilEsc());
        $this->assertInstanceOf(PresentationProperties::class, $object->setLoopContinuouslyUntilEsc());
        $this->assertFalse($object->isLoopContinuouslyUntilEsc());
    }

    public function testLastView(): void
    {
        $object = new PresentationProperties();
        $this->assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
        $this->assertInstanceOf(PresentationProperties::class, $object->setLastView('AAAA'));
        $this->assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
        $this->assertInstanceOf(PresentationProperties::class, $object->setLastView(PresentationProperties::VIEW_OUTLINE));
        $this->assertEquals(PresentationProperties::VIEW_OUTLINE, $object->getLastView());
        $this->assertInstanceOf(PresentationProperties::class, $object->setLastView());
        $this->assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
    }

    public function testMarkAsFinal(): void
    {
        $object = new PresentationProperties();
        $this->assertFalse($object->isMarkedAsFinal());
        $this->assertInstanceOf(PresentationProperties::class, $object->markAsFinal(true));
        $this->assertTrue($object->isMarkedAsFinal());
        $this->assertInstanceOf(PresentationProperties::class, $object->markAsFinal(false));
        $this->assertFalse($object->isMarkedAsFinal());
        $this->assertInstanceOf(PresentationProperties::class, $object->markAsFinal());
        $this->assertTrue($object->isMarkedAsFinal());
    }

    /**
     * @dataProvider dataProviderSlideshowType
     */
    public function testSlideshowType(?string $value, string $expected): void
    {
        $object = new PresentationProperties();
        // Default
        $this->assertEquals(PresentationProperties::SLIDESHOW_TYPE_PRESENT, $object->getSlideshowType());
        // Set value
        if (is_null($value)) {
            $this->assertInstanceOf(PresentationProperties::class, $object->setSlideshowType());
        } else {
            $this->assertInstanceOf(PresentationProperties::class, $object->setSlideshowType($value));
        }
        // Check value
        $this->assertEquals($expected, $object->getSlideshowType());
    }

    /**
     * @return array<array<string|null>>
     */
    public function dataProviderSlideshowType(): array
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
        $this->assertNull($object->getThumbnailPath());
        $this->assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath('AAAA'));
        $this->assertNull($object->getThumbnailPath());
        $this->assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath());
        $this->assertNull($object->getThumbnailPath());
        $this->assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath($imagePath));
        $this->assertEquals($imagePath, $object->getThumbnailPath());
        $this->assertInstanceOf(PresentationProperties::class, $object->setThumbnailPath());
        $this->assertEquals($imagePath, $object->getThumbnailPath());
    }

    public function testZoom(): void
    {
        $object = new PresentationProperties();
        $this->assertEquals(1, $object->getZoom());
        $this->assertInstanceOf(PresentationProperties::class, $object->setZoom(2.3));
        $this->assertEquals(2.3, $object->getZoom());
        $this->assertInstanceOf(PresentationProperties::class, $object->setZoom());
        $this->assertEquals(1, $object->getZoom());
    }
}
