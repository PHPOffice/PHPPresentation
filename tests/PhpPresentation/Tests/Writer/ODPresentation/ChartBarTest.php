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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Writer\ODPresentation;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Charts
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation\Charts
 */
class ChartBarTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }

    public function testChartBar()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));

        $oBar = new Bar();
        $oBar->addSeries($oSeries);

        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:bar', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'chart:vertical', 'Object 1/content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'chart:three-dimensional', 'Object 1/content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'chart:right-angled-axes', 'Object 1/content.xml'));
    }

    public function testChartBarHorizontal()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));

        $oBar = new Bar();
        $oBar->setBarDirection(Bar::DIRECTION_HORIZONTAL);
        $oBar->addSeries($oSeries);

        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:bar', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:vertical', 'Object 1/content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'chart:three-dimensional', 'Object 1/content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'chart:right-angled-axes', 'Object 1/content.xml'));
    }
}
