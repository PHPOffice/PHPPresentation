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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\AbstractGraphic;

/**
 * Test class for Table element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\AbstractGraphic
 */
class AbstractGraphicTest extends \PHPUnit_Framework_TestCase
{
    public function testWidthAndHeight()
    {
        $min = 10;
        $max = 20;
        /** @var AbstractGraphic $stub */
        $stub = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic');
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(false));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($min));
        $this->assertEquals($min, $stub->getWidth());
        $this->assertEquals(0, $stub->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($max));
        $this->assertEquals($min, $stub->getWidth());
        $this->assertEquals($max, $stub->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        $this->assertEquals($min, $stub->getWidth());
        $this->assertEquals($max, $stub->getHeight());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(true));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        $this->assertEquals($max, $stub->getWidth());
        $this->assertEquals($max * ($max / $min), $stub->getHeight());
        $this->assertEquals($max, $stub->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        $this->assertEquals($min * ($max / ($max * ($max / $min))), $stub->getWidth());
        $this->assertEquals($min, $stub->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        $this->assertEquals($min, $stub->getWidth());
        $this->assertEquals($max, $stub->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($max, $min));
        $this->assertEquals($min * ($min / $max), $stub->getWidth());
        $this->assertEquals($min, $stub->getHeight());
    }

    public function testWidthAndHeight2()
    {
        $min = 10;
        $max = 20;
        /** @var AbstractGraphic $stub */
        $stub = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic');
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(false));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        $this->assertEquals($max, $stub->getWidth());
        $this->assertEquals(0, $stub->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        $this->assertEquals($max, $stub->getWidth());
        $this->assertEquals($min, $stub->getHeight());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(true));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        $this->assertEquals($max, $stub->getWidth());
        $this->assertEquals($max * ($min / $max), $stub->getHeight());
        $this->assertEquals($max, $stub->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        $this->assertEquals($max * ($min / ($max * ($min / $max))), $stub->getWidth());
        $this->assertEquals($min, $stub->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        $this->assertEquals($min, $stub->getWidth());
        $this->assertEquals($min * ($min / ($max * ($min / ($max * ($min / $max))))), $stub->getHeight());
    }
}
