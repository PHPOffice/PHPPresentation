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

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Writer\ODPresentation;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 */
class ChartAreaTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }

    public function testChartArea()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->getFill()->setStartColor(new Color('FF93A9CE'));
        
        $oArea = new Area();
        $oArea->addSeries($oSeries);
        
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oArea);
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:area', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:area', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\']/style:graphic-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'draw:fill', 'Object 1/content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:fill-color', 'Object 1/content.xml'));
        $this->assertEquals('#93A9CE', $pres->getElementAttribute($element, 'draw:fill-color', 'Object 1/content.xml'));
    }
}
