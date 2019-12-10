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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Slide\SlideMaster
 */
class SlideMasterTest extends TestCase
{
    public function testBase()
    {
        $object = new SlideMaster();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide', $object);
        static::assertNull($object->getParent());
        static::assertInstanceOf('\\ArrayObject', $object->getShapeCollection());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->colorMap);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->getBackground());
        static::assertEquals('FFFFFF', $object->getBackground()->getColor()->getRGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }

    public function testLayout()
    {
        $object = new SlideMaster();

        // Mock Post
        $mockSlideLayout = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Slide\SlideLayout', array($object));

        static::assertEmpty($object->getAllSlideLayouts());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideLayout', $object->createSlideLayout());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideLayout', $object->addSlideLayout($mockSlideLayout));
        static::assertCount(2, $object->getAllSlideLayouts());
    }

    public function testSchemeColors()
    {
        // Mock Pre
        $mockSchemeColorAccent1 = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Style\SchemeColor');
        $mockSchemeColorAccent1->setValue('accent1');
        $mockSchemeColorAccent1->setRGB('ABCDEF');
        $mockSchemeColorNew = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Style\SchemeColor');
        $mockSchemeColorNew->setValue('new');
        $mockSchemeColorNew->setRGB('ABCDEF');

        $object = new SlideMaster();

        static::assertInternalType('array', $object->getAllSchemeColors());
        static::assertCount(12, $object->getAllSchemeColors());
        // Add idem value
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorAccent1));
        static::assertCount(12, $object->getAllSchemeColors());
        // Add new value
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorNew));
        static::assertCount(13, $object->getAllSchemeColors());
    }

    public function testTextStyles()
    {
        // Mock Pre
        $mockTextStyle = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Style\TextStyle');

        $object = new SlideMaster();

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->setTextStyles($mockTextStyle));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }
}
