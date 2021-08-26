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

        $this->assertInstanceOf(Chart\Title::class, $object->getTitle());
        $this->assertInstanceOf(Chart\Legend::class, $object->getLegend());
        $this->assertInstanceOf(Chart\PlotArea::class, $object->getPlotArea());
        $this->assertInstanceOf(Chart\View3D::class, $object->getView3D());
    }

    public function testClone(): void
    {
        $object = new Chart();

        $oClone = clone $object;

        $this->assertInstanceOf(Chart::class, $oClone);
        $this->assertInstanceOf(Chart\Title::class, $oClone->getTitle());
        $this->assertInstanceOf(Chart\Legend::class, $oClone->getLegend());
        $this->assertInstanceOf(Chart\PlotArea::class, $oClone->getPlotArea());
        $this->assertInstanceOf(Chart\View3D::class, $oClone->getView3D());
    }

    public function testDisplayBlankAs(): void
    {
        $object = new Chart();

        $this->assertEquals(Chart::BLANKAS_ZERO, $object->getDisplayBlankAs());
        $this->assertInstanceOf(Chart::class, $object->setDisplayBlankAs(Chart::BLANKAS_GAP));
        $this->assertEquals(Chart::BLANKAS_GAP, $object->getDisplayBlankAs());
        $this->assertInstanceOf(Chart::class, $object->setDisplayBlankAs(Chart::BLANKAS_ZERO));
        $this->assertEquals(Chart::BLANKAS_ZERO, $object->getDisplayBlankAs());
        $this->assertInstanceOf(Chart::class, $object->setDisplayBlankAs(Chart::BLANKAS_SPAN));
        $this->assertEquals(Chart::BLANKAS_SPAN, $object->getDisplayBlankAs());
        $this->assertInstanceOf(Chart::class, $object->setDisplayBlankAs('Unauthorized value'));
        $this->assertEquals(Chart::BLANKAS_SPAN, $object->getDisplayBlankAs());
    }

    public function testIncludeSpreadsheet(): void
    {
        $object = new Chart();

        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf(Chart::class, $object->setIncludeSpreadsheet());
        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf(Chart::class, $object->setIncludeSpreadsheet(false));
        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf(Chart::class, $object->setIncludeSpreadsheet(true));
        $this->assertTrue($object->hasIncludedSpreadsheet());
    }
}
