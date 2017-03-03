<?php

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\Common\Drawing as CommonDrawing;
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
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\ODPresentation;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation\ObjectsChart
 */
class ObjectsChartTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    public function testAxisFont()
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->getFont()->getColor()->setRGB('AABBCC');
        $oShape->getPlotArea()->getAxisX()->getFont()->setItalic(true);

        $oShape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('00FF00');
        $oShape->getPlotArea()->getAxisY()->getFont()->setSize(16);
        $oShape->getPlotArea()->getAxisY()->getFont()->setName('Arial');

        $this->assertZipFileExists('Object 1/content.xml');

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:color', '#AABBCC');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'italic');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-size', '10pt');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-family', 'Calibri');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:color', '#00FF00');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'normal');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-size', '16pt');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-family', 'Arial');
    }

    public function testLegend()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row/table:table-cell';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row/table:table-cell[@office:value-type=\'string\']';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
    }

    public function testTitleVisibility()
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        $elementTitle = '/office:document-content/office:body/office:chart/chart:chart/chart:title';
        $elementStyle = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleTitle\']';

        $this->assertTrue($oShape->getTitle()->isVisible());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(true));
        $this->assertZipXmlElementExists('Object 1/content.xml', $elementTitle);
        $this->assertZipXmlElementExists('Object 1/content.xml', $elementStyle);

        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(false));
        $this->resetPresentationFile();
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $elementTitle);
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $elementStyle);
    }

    public function testTypeArea()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->getFill()->setStartColor(new Color('FF93A9CE'));
        $oArea = new Area();
        $oArea->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oArea);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:area');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:area');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\']/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'draw:fill');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'draw:fill-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill-color', '#93A9CE');
    }

    public function testTypeAxisBounds()
    {
        $value = rand(0, 100);

        $oSeries = new Series('Downloads', array('A' => 1, 'B' => 2, 'C' => 4, 'D' => 3, 'E' => 2));
        $oSeries->getFill()->setStartColor(new Color('FFAABBCC'));
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:maximum');

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:minimum', $value);

        $oShape->getPlotArea()->getAxisX()->setMinBounds(null);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:maximum', $value);

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:minimum', $value);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:maximum', $value);
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
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'false');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '0');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:percentage');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');
    }

    public function testTypeBarGroupingStacked()
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2)));
        $oBar->setBarGrouping(Bar::GROUPING_STACKED);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '100');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:percentage');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');
    }

    public function testTypeBarGroupingPercentStacked()
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2)));
        $oBar->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '100');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:percentage', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'percentage');
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
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'true');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
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
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar3D);
        
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');
        
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'false');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:three-dimensional', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:right-angled-axes', 'true');
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
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar3D);
        
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:three-dimensional', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:right-angled-axes', 'true');
    }
    
    public function testTypeLine()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);
    
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:line');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-inner', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-outer', 'false');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.026cm');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#878787');


        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-inner', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-outer', 'false');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.026cm');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#878787');
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
            $this->resetPresentationFile();
            $this->oPresentation->removeSlideByIndex(0)->createSlide();

            $expectedSizePts = rand(1, 100);
            $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSizePts), 2, '.', '').'cm';
            $expectedElementGrid = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\''.$arrayTest['dimension'].'\']/chart:grid';
            $expectedElementStyle = '/office:document-content/office:automatic-styles/style:style[@style:name=\''.$arrayTest['styleName'].'\']/style:graphic-properties';

            $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
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

            $this->assertZipFileExists('Object 1/content.xml');
            $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElementGrid);
            $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElementGrid, 'chart:style-name');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElementGrid, 'chart:style-name', $arrayTest['styleName']);
            $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElementGrid, 'chart:class');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElementGrid, 'chart:class', $arrayTest['styleClass']);

            $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElementStyle);
            $this->assertZipXmlAttributeStartsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-width', $expectedSizeCm);
            $this->assertZipXmlAttributeEndsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-width', 'cm');
            $this->assertZipXmlAttributeStartsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-color', '#');
            $this->assertZipXmlAttributeEndsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-color', $expectedColor->getRGB());
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

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
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

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', $expectedSymbol1);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-width', $expectedSizeCm);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-height', $expectedSizeCm);

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oLine->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'horizontal-bar');

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oLine->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'circle');

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oLine->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'arrow-up');

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oLine->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-name');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-width');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-height');
    }

    public function testTypeLineSeriesOutline()
    {
        $expectedWidth = rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pointsToCentimeters($expectedWidth), 3, '.', '').'cm';

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:graphic-properties';

        $oColor = new Color(Color::COLOR_YELLOW);
        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor($oColor);
        $oOutline->setWidth($expectedWidth); // (in points)

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
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

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');

        $oSeries->setOutline($oOutline);
        $oLine->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $oColor->getRGB());
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
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie);
    
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:circle');
        
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');
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
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie3D);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:circle');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');
    }
    
    public function testTypePie3DExplosion()
    {
        $value = rand(0, 100);
        
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oPie3D = new Pie3D();
        $oPie3D->setExplosion($value);
        $oPie3D->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie3D);
    
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:pie-offset', $value);
    }
    
    public function testTypeScatter()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oSeries->setShowSeriesName(true);
        $oScatter = new Scatter();
        $oScatter->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oScatter);
    
        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:scatter');
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

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
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

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', $expectedSymbol1);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-width', $expectedSizeCm);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-height', $expectedSizeCm);

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oScatter->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'horizontal-bar');

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oScatter->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'circle');

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oScatter->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'arrow-up');

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oScatter->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-name');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-width');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-height');
    }

    public function testTypeScatterSeriesOutline()
    {
        $expectedWidth = rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pointsToCentimeters($expectedWidth), 3, '.', '').'cm';

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:graphic-properties';

        $oColor = new Color(Color::COLOR_YELLOW);
        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($oColor);
        $oOutline->setWidth($expectedWidth); // (in points)

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
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

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');

        $oSeries->setOutline($oOutline);
        $oScatter->setSeries(array($oSeries));
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $oColor->getRGB());
    }

    public function testSeries()
    {
        $oSeries = new Series('Series', array('Jan' => 1, 'Feb' => 5, 'Mar' => 2));
        $oPie = new Pie();
        $oPie->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\']/style:chart-properties';

        // $showCategoryName = false / $showPercentage = false / $showValue = true
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');

        $oSeries->setShowValue(false);
        $this->resetPresentationFile();

        // $showCategoryName = false / $showPercentage = false / $showValue = false
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');

        // $showCategoryName = false / $showPercentage = true / $showValue = true
        $oSeries->setShowValue(true);
        $oSeries->setShowPercentage(true);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value-and-percentage');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');

        // $showCategoryName = false / $showPercentage = true / $showValue = false
        $oSeries->setShowValue(false);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'percentage');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');

        // $showCategoryName = false / $showPercentage = true / $showValue = false
        $oSeries->setShowCategoryName(true);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-text');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-text', 'true');
    }
}
