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

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Writer\ODPresentation;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation\ObjectsChart
 */
class ObjectsChartTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }

    public function testAxisFont()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->getFont()->getColor()->setRGB('AABBCC');
        $oShape->getPlotArea()->getAxisX()->getFont()->setItalic(true);

        $oShape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('00FF00');
        $oShape->getPlotArea()->getAxisY()->getFont()->setSize(16);
        $oShape->getPlotArea()->getAxisY()->getFont()->setName('Arial');

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:bar', $oXMLDoc->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:text-properties';
        $this->assertTrue($oXMLDoc->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('#AABBCC', $oXMLDoc->getElementAttribute($element, 'fo:color', 'Object 1/content.xml'));//Color XAxis
        $this->assertEquals('italic', $oXMLDoc->getElementAttribute($element, 'fo:font-style', 'Object 1/content.xml'));//Italic XAxis
        $this->assertEquals('10pt', $oXMLDoc->getElementAttribute($element, 'fo:font-size', 'Object 1/content.xml'));//Size XAxis
        $this->assertEquals('Calibri', $oXMLDoc->getElementAttribute($element, 'fo:font-family', 'Object 1/content.xml'));//Size XAxis

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:text-properties';
        $this->assertTrue($oXMLDoc->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('#00FF00', $oXMLDoc->getElementAttribute($element, 'fo:color', 'Object 1/content.xml'));//Color YAxis
        $this->assertEquals('normal', $oXMLDoc->getElementAttribute($element, 'fo:font-style', 'Object 1/content.xml'));//Italic YAxis
        $this->assertEquals('16pt', $oXMLDoc->getElementAttribute($element, 'fo:font-size', 'Object 1/content.xml'));//Size YAxis
        $this->assertEquals('Arial', $oXMLDoc->getElementAttribute($element, 'fo:font-family', 'Object 1/content.xml'));//Size YAxis

    }

    public function testLegend()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row/table:table-cell';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row/table:table-cell[@office:value-type=\'string\']';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
    }

    public function testTitleVisibility()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        $elementTitle = '/office:document-content/office:body/office:chart/chart:chart/chart:title';
        $elementStyle = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleTitle\']';

        $this->assertTrue($oShape->getTitle()->isVisible());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(true));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->elementExists($elementTitle, 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->elementExists($elementStyle, 'Object 1/content.xml'));

        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(false));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertFalse($oXMLDoc->elementExists($elementTitle, 'Object 1/content.xml'));
        $this->assertFalse($oXMLDoc->elementExists($elementStyle, 'Object 1/content.xml'));
    }

    public function testTypeArea()
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

    public function testTypeBar()
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
        $this->assertEquals('false', $pres->getElementAttribute($element, 'chart:stacked', 'Object 1/content.xml'));
        $this->assertEquals('0', $pres->getElementAttribute($element, 'chart:overlap', 'Object 1/content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'chart:percentage', 'Object 1/content.xml'));
        $this->assertEquals('value', $pres->getElementAttribute($element, 'chart:data-label-number', 'Object 1/content.xml'));
    }

    public function testTypeBarGroupingStacked()
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2)));
        $oBar->setBarGrouping(Bar::GROUPING_STACKED);

        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:stacked', 'Object 1/content.xml'));
        $this->assertEquals('100', $pres->getElementAttribute($element, 'chart:overlap', 'Object 1/content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'chart:percentage', 'Object 1/content.xml'));
        $this->assertEquals('value', $pres->getElementAttribute($element, 'chart:data-label-number', 'Object 1/content.xml'));
    }

    public function testTypeBarGroupingPercentStacked()
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2)));
        $oBar->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);

        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:stacked', 'Object 1/content.xml'));
        $this->assertEquals('100', $pres->getElementAttribute($element, 'chart:overlap', 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:percentage', 'Object 1/content.xml'));
        $this->assertEquals('percentage', $pres->getElementAttribute($element, 'chart:data-label-number', 'Object 1/content.xml'));
    }

    public function testTypeBarHorizontal()
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

    public function testTypeBar3D()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));

        $oBar3D = new Bar3D();
        $oBar3D->addSeries($oSeries);
        
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oBar3D);
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:bar', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'chart:vertical', 'Object 1/content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'chart:three-dimensional', 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:three-dimensional', 'Object 1/content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'chart:right-angled-axes', 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:right-angled-axes', 'Object 1/content.xml'));
    }

    public function testTypeBar3DHorizontal()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));

        $oBar3D = new Bar3D();
        $oBar3D->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
        $oBar3D->addSeries($oSeries);
        
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oBar3D);
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:bar', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:vertical', 'Object 1/content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'chart:three-dimensional', 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:three-dimensional', 'Object 1/content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'chart:right-angled-axes', 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:right-angled-axes', 'Object 1/content.xml'));
    }
    
    public function testTypeLine()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
    
        $oLine = new Line();
        $oLine->addSeries($oSeries);
    
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oLine);
    
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
    
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:line', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'chart:tick-marks-major-inner', 'Object 1/content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'chart:tick-marks-major-outer', 'Object 1/content.xml'));
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:graphic-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('0.026cm', $pres->getElementAttribute($element, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertEquals('#878787', $pres->getElementAttribute($element, 'svg:stroke-color', 'Object 1/content.xml'));
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'chart:tick-marks-major-inner', 'Object 1/content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'chart:tick-marks-major-outer', 'Object 1/content.xml'));
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:graphic-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('0.026cm', $pres->getElementAttribute($element, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertEquals('#878787', $pres->getElementAttribute($element, 'svg:stroke-color', 'Object 1/content.xml'));
    }

    public function testTypeLineGridlines()
    {
        $arrayTests = array(
            array(
                'dimension' => 'x',
                'styleName' => 'styleAxisXGridlinesMajor',
                'styleClass' => 'major',
                'methodAxis' => 'getAxisX',
                'methodGrid' => 'setMajorGridlines'
            ),
            array(
                'dimension' => 'x',
                'styleName' => 'styleAxisXGridlinesMinor',
                'styleClass' => 'minor',
                'methodAxis' => 'getAxisX',
                'methodGrid' => 'setMinorGridlines'
            ),
            array(
                'dimension' => 'y',
                'styleName' => 'styleAxisYGridlinesMajor',
                'styleClass' => 'major',
                'methodAxis' => 'getAxisY',
                'methodGrid' => 'setMajorGridlines'
            ),
            array(
                'dimension' => 'y',
                'styleName' => 'styleAxisYGridlinesMinor',
                'styleClass' => 'minor',
                'methodAxis' => 'getAxisY',
                'methodGrid' => 'setMinorGridlines'
            ),
        );
        $expectedColor = new Color(Color::COLOR_BLUE);
        foreach ($arrayTests as $arrayTest) {
            $expectedSizePts = rand(1, 100);
            $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSizePts), 2, '.', '').'cm';
            $expectedElementGrid = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\''.$arrayTest['dimension'].'\']/chart:grid';
            $expectedElementStyle = '/office:document-content/office:automatic-styles/style:style[@style:name=\''.$arrayTest['styleName'].'\']/style:graphic-properties';

            $oPresentation = new PhpPresentation();
            $oSlide = $oPresentation->getActiveSlide();
            $oShape = $oSlide->createChartShape();
            $oLine = new Line();
            $oLine->addSeries(new Series('Downloads', array(
                'A' => 1,
                'B' => 2,
                'C' => 4,
                'D' => 3,
                'E' => 2,
            )));
            $oShape->getPlotArea()->setType($oLine);
            $oGridlines = new Gridlines();
            $oGridlines->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($expectedColor);
            $oGridlines->getOutline()->setWidth($expectedSizePts);
            $oShape->getPlotArea()->{$arrayTest['methodAxis']}()->{$arrayTest['methodGrid']}($oGridlines);

            $oXMLDoc = TestHelperDOCX::getDocument($oPresentation, 'ODPresentation');
            $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));
            $this->assertTrue($oXMLDoc->elementExists($expectedElementGrid, 'Object 1/content.xml'));
            $this->assertTrue($oXMLDoc->attributeElementExists($expectedElementGrid, 'chart:style-name', 'Object 1/content.xml'));
            $this->assertEquals($arrayTest['styleName'], $oXMLDoc->getElementAttribute($expectedElementGrid, 'chart:style-name', 'Object 1/content.xml'));
            $this->assertTrue($oXMLDoc->attributeElementExists($expectedElementGrid, 'chart:class', 'Object 1/content.xml'));
            $this->assertEquals($arrayTest['styleClass'], $oXMLDoc->getElementAttribute($expectedElementGrid, 'chart:class', 'Object 1/content.xml'));
            $this->assertTrue($oXMLDoc->elementExists($expectedElementStyle, 'Object 1/content.xml'));
            $this->assertStringStartsWith($expectedSizeCm, $oXMLDoc->getElementAttribute($expectedElementStyle, 'svg:stroke-width', 'Object 1/content.xml'));
            $this->assertStringEndsWith('cm', $oXMLDoc->getElementAttribute($expectedElementStyle, 'svg:stroke-width', 'Object 1/content.xml'));
            $this->assertStringStartsWith('#', $oXMLDoc->getElementAttribute($expectedElementStyle, 'svg:stroke-color', 'Object 1/content.xml'));
            $this->assertStringEndsWith($expectedColor->getRGB(), $oXMLDoc->getElementAttribute($expectedElementStyle, 'svg:stroke-color', 'Object 1/content.xml'));

            unset($oPresentation);
            TestHelperDOCX::clear();
        }
    }

    public function testTypeLineMarker()
    {
        $expectedSymbol1 = Marker::SYMBOL_PLUS;
        $expectedSymbol2 = Marker::SYMBOL_DASH;
        $expectedSymbol3 = Marker::SYMBOL_DOT;
        $expectedSymbol4 = Marker::SYMBOL_TRIANGLE;
        $expectedSymbol5 = Marker::SYMBOL_NONE;

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:chart-properties';
        $expectedSize = rand(1, 100);
        $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSize), 2, '.', '').'cm';

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oSeries = new Series('Downloads', array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        ));
        $oSeries->getMarker()->setSymbol($expectedSymbol1)->setSize($expectedSize);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'Object 1/content.xml'));
        $this->assertEquals($expectedSymbol1, $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));
        $this->assertEquals($expectedSizeCm, $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-width', 'Object 1/content.xml'));
        $this->assertEquals($expectedSizeCm, $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-height', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oLine->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertEquals('horizontal-bar', $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oLine->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertEquals('circle', $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oLine->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertEquals('arrow-up', $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oLine->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertFalse($oXMLDoc->attributeElementExists($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));
        $this->assertFalse($oXMLDoc->attributeElementExists($expectedElement, 'chart:symbol-width', 'Object 1/content.xml'));
        $this->assertFalse($oXMLDoc->attributeElementExists($expectedElement, 'chart:symbol-height', 'Object 1/content.xml'));
    }

    public function testTypeLineSeriesOutline()
    {
        $expectedWidth = rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pointsToCentimeters($expectedWidth), 3, '.', '').'cm';

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:graphic-properties';

        $oColor = new Color(Color::COLOR_YELLOW);
        $oOutline = new Outline();
        // Define the color
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor($oColor);
        // Define the width (in points)
        $oOutline->setWidth($expectedWidth);

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        ));
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertEquals('0.079cm', $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));
        $this->assertEquals('#4a7ebb', $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));

        $oSeries->setOutline($oOutline);
        $oLine->setSeries(array($oSeries));

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertEquals($expectedWidthCm, $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));
        $this->assertEquals('#'.$oColor->getRGB(), $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));
    }
    
    public function testTypePie()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
    
        $oPie = new Pie();
        $oPie->addSeries($oSeries);
    
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oPie);
    
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
    
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:circle', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:reverse-direction', 'Object 1/content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:reverse-direction', 'Object 1/content.xml'));
    }

    public function testTypePie3D()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));

        $oPie3D = new Pie3D();
        $oPie3D->addSeries($oSeries);

        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oPie3D);

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:circle', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:reverse-direction', 'Object 1/content.xml'));

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'chart:reverse-direction', 'Object 1/content.xml'));
    }
    
    public function testTypePie3DExplosion()
    {
        $value = rand(0, 100);
        
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
    
        $oPie3D = new Pie3D();
        $oPie3D->setExplosion($value);
        $oPie3D->addSeries($oSeries);
    
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oPie3D);
    
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
    
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:chart-properties';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals($value, $pres->getElementAttribute($element, 'chart:pie-offset', 'Object 1/content.xml'));
    }
    
    public function testTypeScatter()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
    
        $oScatter = new Scatter();
        $oScatter->addSeries($oSeries);
    
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType($oScatter);
    
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
    
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertTrue($pres->elementExists($element, 'Object 1/content.xml'));
        $this->assertEquals('chart:scatter', $pres->getElementAttribute($element, 'chart:class', 'Object 1/content.xml'));
    }

    public function testTypeScatterMarker()
    {
        $expectedSymbol1 = Marker::SYMBOL_PLUS;
        $expectedSymbol2 = Marker::SYMBOL_DASH;
        $expectedSymbol3 = Marker::SYMBOL_DOT;
        $expectedSymbol4 = Marker::SYMBOL_TRIANGLE;
        $expectedSymbol5 = Marker::SYMBOL_NONE;

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:chart-properties';
        $expectedSize = rand(1, 100);
        $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSize), 2, '.', '').'cm';

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        ));
        $oSeries->getMarker()->setSymbol($expectedSymbol1)->setSize($expectedSize);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'Object 1/content.xml'));
        $this->assertEquals($expectedSymbol1, $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));
        $this->assertEquals($expectedSizeCm, $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-width', 'Object 1/content.xml'));
        $this->assertEquals($expectedSizeCm, $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-height', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oScatter->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertEquals('horizontal-bar', $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oScatter->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertEquals('circle', $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oScatter->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertEquals('arrow-up', $oXMLDoc->getElementAttribute($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oScatter->setSeries(array($oSeries));
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertFalse($oXMLDoc->attributeElementExists($expectedElement, 'chart:symbol-name', 'Object 1/content.xml'));
        $this->assertFalse($oXMLDoc->attributeElementExists($expectedElement, 'chart:symbol-width', 'Object 1/content.xml'));
        $this->assertFalse($oXMLDoc->attributeElementExists($expectedElement, 'chart:symbol-height', 'Object 1/content.xml'));
    }

    public function testTypeScatterSeriesOutline()
    {
        $expectedWidth = rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pointsToCentimeters($expectedWidth), 3, '.', '').'cm';

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:graphic-properties';

        $oColor = new Color(Color::COLOR_YELLOW);
        $oOutline = new Outline();
        // Define the color
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor($oColor);
        // Define the width (in points)
        $oOutline->setWidth($expectedWidth);

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        ));
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertEquals('0.079cm', $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));
        $this->assertEquals('#4a7ebb', $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));

        $oSeries->setOutline($oOutline);
        $oScatter->setSeries(array($oSeries));

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $this->assertTrue($oXMLDoc->fileExists('Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertEquals($expectedWidthCm, $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-width', 'Object 1/content.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));
        $this->assertEquals('#'.$oColor->getRGB(), $oXMLDoc->getElementAttribute($expectedElement, 'svg:stroke-color', 'Object 1/content.xml'));
    }
}
