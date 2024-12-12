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

use PhpOffice\PhpPresentation\Shape\Chart;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Chart element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart
 */
class ChartTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Chart();

        self::assertInstanceOf(Chart\Title::class, $object->getTitle());
        self::assertInstanceOf(Chart\Legend::class, $object->getLegend());
        self::assertInstanceOf(Chart\PlotArea::class, $object->getPlotArea());
        self::assertInstanceOf(Chart\View3D::class, $object->getView3D());
    }

    public function testClone(): void
    {
        $object = new Chart();

        $oClone = clone $object;

        self::assertInstanceOf(Chart::class, $oClone);
        self::assertInstanceOf(Chart\Title::class, $oClone->getTitle());
        self::assertInstanceOf(Chart\Legend::class, $oClone->getLegend());
        self::assertInstanceOf(Chart\PlotArea::class, $oClone->getPlotArea());
        self::assertInstanceOf(Chart\View3D::class, $oClone->getView3D());
    }

    public function testDisplayBlankAs(): void
    {
        $object = new Chart();

        self::assertEquals(Chart::BLANKAS_ZERO, $object->getDisplayBlankAs());
        self::assertInstanceOf(Chart::class, $object->setDisplayBlankAs(Chart::BLANKAS_GAP));
        self::assertEquals(Chart::BLANKAS_GAP, $object->getDisplayBlankAs());
        self::assertInstanceOf(Chart::class, $object->setDisplayBlankAs(Chart::BLANKAS_ZERO));
        self::assertEquals(Chart::BLANKAS_ZERO, $object->getDisplayBlankAs());
        self::assertInstanceOf(Chart::class, $object->setDisplayBlankAs(Chart::BLANKAS_SPAN));
        self::assertEquals(Chart::BLANKAS_SPAN, $object->getDisplayBlankAs());
        self::assertInstanceOf(Chart::class, $object->setDisplayBlankAs('Unauthorized value'));
        self::assertEquals(Chart::BLANKAS_SPAN, $object->getDisplayBlankAs());
    }

    public function testIncludeSpreadsheet(): void
    {
        $object = new Chart();

        self::assertFalse($object->hasIncludedSpreadsheet());
        self::assertInstanceOf(Chart::class, $object->setIncludeSpreadsheet());
        self::assertFalse($object->hasIncludedSpreadsheet());
        self::assertInstanceOf(Chart::class, $object->setIncludeSpreadsheet(false));
        self::assertFalse($object->hasIncludedSpreadsheet());
        self::assertInstanceOf(Chart::class, $object->setIncludeSpreadsheet(true));
        self::assertTrue($object->hasIncludedSpreadsheet());
    }
}
