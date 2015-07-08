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

use PhpOffice\PhpPresentation\Shape\Chart;

/**
 * Test class for Chart element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart
 */
class ChartTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Chart();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->getTitle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->getLegend());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->getPlotArea());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->getView3D());
    }
    

    public function testClone()
    {
        $object = new Chart();
        
        $oClone = clone $object;
    
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart', $oClone);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $oClone->getTitle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $oClone->getLegend());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $oClone->getPlotArea());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $oClone->getView3D());
    }

    public function testIncludeSpreadsheet()
    {
        $object = new Chart();

        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart', $object->setIncludeSpreadsheet());
        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart', $object->setIncludeSpreadsheet(false));
        $this->assertFalse($object->hasIncludedSpreadsheet());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart', $object->setIncludeSpreadsheet(true));
        $this->assertTrue($object->hasIncludedSpreadsheet());
    }
}
