<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shape;

use PhpOffice\PhpPowerpoint\Shape\Chart;

/**
 * Test class for Chart element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Chart
 */
class ChartTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Chart();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Title', $object->getTitle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Legend', $object->getLegend());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->getPlotArea());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\View3D', $object->getView3D());
    }

    public function testIncludeSpreadsheet()
    {
        $object = new Chart();

        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart', $object->setIncludeSpreadsheet());
        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart', $object->setIncludeSpreadsheet(false));
        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart', $object->setIncludeSpreadsheet(true));
        $this->assertTrue($object->hasIncludedSpreadsheet());
    }
}
