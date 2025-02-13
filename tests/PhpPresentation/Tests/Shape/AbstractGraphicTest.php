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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\AbstractGraphic;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Table element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\AbstractGraphic
 */
class AbstractGraphicTest extends TestCase
{
    public function testWidthAndHeight(): void
    {
        $min = 10;
        $max = 20;
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractGraphic $stub */
            $stub = $this->getMockForAbstractClass(AbstractGraphic::class);
        } else {
            /** @var AbstractGraphic $stub */
            $stub = new class() extends AbstractGraphic {
            };
        }
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(false));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($min));
        self::assertEquals($min, $stub->getWidth());
        self::assertEquals(0, $stub->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($max));
        self::assertEquals($min, $stub->getWidth());
        self::assertEquals($max, $stub->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        self::assertEquals($min, $stub->getWidth());
        self::assertEquals($max, $stub->getHeight());

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(true));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        self::assertEquals($max, $stub->getWidth());
        self::assertEquals($max * ($max / $min), $stub->getHeight());
        self::assertEquals($max, $stub->getWidth());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        self::assertEquals($min * ($max / ($max * ($max / $min))), $stub->getWidth());
        self::assertEquals($min, $stub->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        self::assertEquals($min, $stub->getWidth());
        self::assertEquals($max, $stub->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($max, $min));
        self::assertEquals($min * ($min / $max), $stub->getWidth());
        self::assertEquals($min, $stub->getHeight());
    }

    public function testWidthAndHeight2(): void
    {
        $min = 10;
        $max = 20;
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractGraphic $stub */
            $stub = $this->getMockForAbstractClass(AbstractGraphic::class);
        } else {
            /** @var AbstractGraphic $stub */
            $stub = new class() extends AbstractGraphic {
            };
        }
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(false));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        self::assertEquals($max, $stub->getWidth());
        self::assertEquals(0, $stub->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        self::assertEquals($max, $stub->getWidth());
        self::assertEquals($min, $stub->getHeight());

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setResizeProportional(true));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidth($max));
        self::assertEquals($max, $stub->getWidth());
        self::assertEquals($max * ($min / $max), $stub->getHeight());
        self::assertEquals($max, $stub->getWidth());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setHeight($min));
        self::assertEquals($max * ($min / ($max * ($min / $max))), $stub->getWidth());
        self::assertEquals($min, $stub->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\AbstractGraphic', $stub->setWidthAndHeight($min, $max));
        self::assertEquals($min, $stub->getWidth());
        self::assertEquals($min * ($min / ($max * ($min / ($max * ($min / $max))))), $stub->getHeight());
    }
}
