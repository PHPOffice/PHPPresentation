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

use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Slide\SlideLayout
 */
class SlideLayoutTest extends TestCase
{
    public function testBase(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var SlideMaster $mockSlideMaster */
            $mockSlideMaster = $this->getMockForAbstractClass(SlideMaster::class);
        } else {
            /** @var SlideMaster $mockSlideMaster */
            $mockSlideMaster = new class() extends SlideMaster {
            };
        }

        $object = new SlideLayout($mockSlideMaster);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide', $object);
        self::assertIsArray($object->getShapeCollection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->colorMap);
    }

    public function testLayoutName(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var SlideMaster $mockSlideMaster */
            $mockSlideMaster = $this->getMockForAbstractClass(SlideMaster::class);
        } else {
            /** @var SlideMaster $mockSlideMaster */
            $mockSlideMaster = new class() extends SlideMaster {
            };
        }

        // Expected
        $expectedLayoutName = 'Title' . mt_rand(1, 100);

        $object = new SlideLayout($mockSlideMaster);

        self::assertNull($object->getLayoutName());
        self::assertInstanceOf(SlideLayout::class, $object->setLayoutName($expectedLayoutName));
        self::assertEquals($expectedLayoutName, $object->getLayoutName());
    }

    public function testSlideMaster(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var SlideMaster $mockSlideMaster */
            $mockSlideMaster = $this->getMockForAbstractClass(SlideMaster::class);
        } else {
            /** @var SlideMaster $mockSlideMaster */
            $mockSlideMaster = new class() extends SlideMaster {
            };
        }

        $object = new SlideLayout($mockSlideMaster);

        self::assertInstanceOf(SlideMaster::class, $object->getSlideMaster());
    }
}
