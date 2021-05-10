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
use PHPUnit\Framework\TestCase;

/**
 * Test class for Table element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\AbstractGraphic
 */
class AbstractGraphicTest extends TestCase
{
    public function testWidthAndHeight()
    {
        $min = 10;
        $max = 20;
        /** @var AbstractGraphic $stub */
        $stub = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic');
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(false));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($min));
        static::assertEquals($min, $stub->getWidth());
        static::assertEquals(0, $stub->getHeight());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($max));
        static::assertEquals($min, $stub->getWidth());
        static::assertEquals($max, $stub->getHeight());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        static::assertEquals($min, $stub->getWidth());
        static::assertEquals($max, $stub->getHeight());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(true));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        static::assertEquals($max, $stub->getWidth());
        static::assertEquals($max * ($max / $min), $stub->getHeight());
        static::assertEquals($max, $stub->getWidth());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        static::assertEquals($min * ($max / ($max * ($max / $min))), $stub->getWidth());
        static::assertEquals($min, $stub->getHeight());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        static::assertEquals($min, $stub->getWidth());
        static::assertEquals($max, $stub->getHeight());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($max, $min));
        static::assertEquals($min * ($min / $max), $stub->getWidth());
        static::assertEquals($min, $stub->getHeight());
    }

    public function testWidthAndHeight2()
    {
        $min = 10;
        $max = 20;
        /** @var AbstractGraphic $stub */
        $stub = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic');
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(false));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        static::assertEquals($max, $stub->getWidth());
        static::assertEquals(0, $stub->getHeight());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        static::assertEquals($max, $stub->getWidth());
        static::assertEquals($min, $stub->getHeight());

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(true));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        static::assertEquals($max, $stub->getWidth());
        static::assertEquals($max * ($min / $max), $stub->getHeight());
        static::assertEquals($max, $stub->getWidth());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        static::assertEquals($max * ($min / ($max * ($min / $max))), $stub->getWidth());
        static::assertEquals($min, $stub->getHeight());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        static::assertEquals($min, $stub->getWidth());
        static::assertEquals($min * ($min / ($max * ($min / ($max * ($min / $max))))), $stub->getHeight());
    }
}
