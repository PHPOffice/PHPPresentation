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
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide', $object);
        $this->assertNull($object->getParent());
        $this->assertInstanceOf('\\ArrayObject', $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->colorMap);
        /** @var Color $background */
        $background = $object->getBackground();
        $this->assertInstanceOf(Color::class, $background);
        $this->assertEquals('FFFFFF', $background->getColor()->getRGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }

    public function testLayout(): void
    {
        $object = new SlideMaster();

        // Mock Post
        /** @var SlideLayout $mockSlideLayout */
        $mockSlideLayout = $this->getMockForAbstractClass(SlideLayout::class, [$object]);

        $this->assertEmpty($object->getAllSlideLayouts());
        $this->assertInstanceOf(SlideLayout::class, $object->createSlideLayout());
        $this->assertInstanceOf(SlideLayout::class, $object->addSlideLayout($mockSlideLayout));
        $this->assertCount(2, $object->getAllSlideLayouts());
    }

    public function testSchemeColors(): void
    {
        // Mock Pre
        /** @var SchemeColor $mockSchemeColorAccent1 */
        $mockSchemeColorAccent1 = $this->getMockForAbstractClass(SchemeColor::class);
        $mockSchemeColorAccent1->setValue('accent1');
        $mockSchemeColorAccent1->setRGB('ABCDEF');
        /** @var SchemeColor $mockSchemeColorNew */
        $mockSchemeColorNew = $this->getMockForAbstractClass(SchemeColor::class);
        $mockSchemeColorNew->setValue('new');
        $mockSchemeColorNew->setRGB('ABCDEF');

        $object = new SlideMaster();

        $this->assertIsArray($object->getAllSchemeColors());
        $this->assertCount(12, $object->getAllSchemeColors());
        // Add idem value
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorAccent1));
        $this->assertCount(12, $object->getAllSchemeColors());
        // Add new value
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorNew));
        $this->assertCount(13, $object->getAllSchemeColors());
    }

    public function testTextStyles(): void
    {
        // Mock Pre
        /** @var TextStyle $mockTextStyle */
        $mockTextStyle = $this->getMockForAbstractClass(TextStyle::class);

        $object = new SlideMaster();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->setTextStyles($mockTextStyle));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }
}
