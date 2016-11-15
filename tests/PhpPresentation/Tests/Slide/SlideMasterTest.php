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

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Slide\SlideMaster
 */
class SlideMasterTest extends \PHPUnit_Framework_TestCase
{
    public function testBase()
    {
        $object = new SlideMaster();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide', $object);
        $this->assertNull($object->getParent());
        $this->assertInstanceOf('\\ArrayObject', $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->colorMap);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->getBackground());
        $this->assertEquals('FFFFFF', $object->getBackground()->getColor()->getRGB());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }

    public function testLayout()
    {
        $object = new SlideMaster();

        // Mock Post
        $mockSlideLayout = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Slide\SlideLayout', array($object));

        $this->assertEmpty($object->getAllSlideLayouts());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideLayout', $object->createSlideLayout());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideLayout', $object->addSlideLayout($mockSlideLayout));
        $this->assertCount(2, $object->getAllSlideLayouts());
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

        $this->assertInternalType('array', $object->getAllSchemeColors());
        $this->assertCount(12, $object->getAllSchemeColors());
        // Add idem value
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorAccent1));
        $this->assertCount(12, $object->getAllSchemeColors());
        // Add new value
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->addSchemeColor($mockSchemeColorNew));
        $this->assertCount(13, $object->getAllSchemeColors());
    }

    public function testTextStyles()
    {
        // Mock Pre
        $mockTextStyle = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Style\TextStyle');

        $object = new SlideMaster();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\SlideMaster', $object->setTextStyles($mockTextStyle));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\TextStyle', $object->getTextStyles());
    }
}
