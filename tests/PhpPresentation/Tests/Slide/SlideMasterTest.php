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

use PhpOffice\PhpPresentation\Slide\Background\Color;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Style\SchemeColor;
use PhpOffice\PhpPresentation\Style\TextStyle;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Slide\SlideMaster
 */
class SlideMasterTest extends TestCase
{
    public function testBase(): void
    {
        $object = new SlideMaster();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide', $object);
        self::assertNull($object->getParent());
        self::assertIsArray($object->getShapeCollection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->colorMap);
        /** @var Color $background */
        $background = $object->getBackground();
        self::assertInstanceOf(Color::class, $background);
        self::assertEquals('FFFFFF', $background->getColor()->getRGB());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }

    public function testLayout(): void
    {
        $object = new SlideMaster();

        // Mock Post
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var SlideLayout $mockSlideLayout */
            $mockSlideLayout = $this->getMockForAbstractClass(SlideLayout::class, [$object]);
        } else {
            /** @var SlideLayout $mockSlideLayout */
            $mockSlideLayout = new class($object) extends SlideLayout {
            };
        }

        self::assertEmpty($object->getAllSlideLayouts());
        self::assertInstanceOf(SlideLayout::class, $object->createSlideLayout());
        self::assertInstanceOf(SlideLayout::class, $object->addSlideLayout($mockSlideLayout));
        self::assertCount(2, $object->getAllSlideLayouts());
    }

    public function testSchemeColors(): void
    {
        // Mock Pre
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var SchemeColor $mockSchemeColorAccent1 */
            $mockSchemeColorAccent1 = $this->getMockForAbstractClass(SchemeColor::class);
        } else {
            /** @var SchemeColor $mockSchemeColorAccent1 */
            $mockSchemeColorAccent1 = new class() extends SchemeColor {
            };
        }
        $mockSchemeColorAccent1->setValue('accent1');
        $mockSchemeColorAccent1->setRGB('ABCDEF');
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var SchemeColor $mockSchemeColorNew */
            $mockSchemeColorNew = $this->getMockForAbstractClass(SchemeColor::class);
        } else {
            /** @var SchemeColor $mockSchemeColorNew */
            $mockSchemeColorNew = new class() extends SchemeColor {
            };
        }
        $mockSchemeColorNew->setValue('new');
        $mockSchemeColorNew->setRGB('ABCDEF');

        $object = new SlideMaster();

        self::assertIsArray($object->getAllSchemeColors());
        self::assertCount(12, $object->getAllSchemeColors());
        // Add idem value
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorAccent1));
        self::assertCount(12, $object->getAllSchemeColors());
        // Add new value
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorNew));
        self::assertCount(13, $object->getAllSchemeColors());
    }

    public function testTextStyles(): void
    {
        // Mock Pre
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var TextStyle $mockTextStyle */
            $mockTextStyle = $this->getMockForAbstractClass(TextStyle::class);
        } else {
            /** @var TextStyle $mockTextStyle */
            $mockTextStyle = new class() extends TextStyle {
            };
        }

        $object = new SlideMaster();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->setTextStyles($mockTextStyle));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }
}
