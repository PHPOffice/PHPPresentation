<?php
/**
 * Created by PhpStorm.
 * User: lefevre_f
 * Date: 01/03/2016
 * Time: 12:35
 */

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

class PptChartsTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var PhpPresentation;
     */
    protected $oPresentation;

    /**
     * @var array
     */
    protected $seriesData = array(
        'A' => 1,
        'B' => 2,
        'C' => 4,
        'D' => 3,
        'E' => 2,
    );

    public function setUp()
    {
        $this->oPresentation = new PhpPresentation();
    }

    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        $this->oPresentation = null;
        TestHelperDOCX::clear();
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage The chart type provided could not be rendered.
     */
    public function testPlotAreaBadType()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $stub = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType');
        $oShape->getPlotArea()->setType($stub);

        TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
    }

    public function testTitleVisibility()
    {
        $element = '/c:chartSpace/c:chart/c:autoTitleDeleted';
        
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertTrue($oShape->getTitle()->isVisible());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(true));
        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('0', $oXMLDoc->getElementAttribute($element, 'val', 'ppt/charts/'.$oShape->getIndexedFilename()));

        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(false));
        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('1', $oXMLDoc->getElementAttribute($element, 'val', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testAxisFont()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oBar = new Bar();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->getFont()->getColor()->setRGB('AABBCC');
        $oShape->getPlotArea()->getAxisX()->getFont()->setItalic(true);
        $oShape->getPlotArea()->getAxisX()->getFont()->setStrikethrough(true);
        $oShape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('00FF00');
        $oShape->getPlotArea()->getAxisY()->getFont()->setSize(16);
        $oShape->getPlotArea()->getAxisY()->getFont()->setUnderline(Font::UNDERLINE_DASH);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill/a:srgbClr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('AABBCC', $oXMLDoc->getElementAttribute($element, 'val', 'ppt/charts/'.$oShape->getIndexedFilename()));  //Color XAxis
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('true', $oXMLDoc->getElementAttribute($element, 'i', 'ppt/charts/'.$oShape->getIndexedFilename())); //Italic XAxis
        $this->assertEquals(10*100, $oXMLDoc->getElementAttribute($element, 'sz', 'ppt/charts/'.$oShape->getIndexedFilename())); //Size XAxis
        $this->assertEquals('sngStrike', $oXMLDoc->getElementAttribute($element, 'strike', 'ppt/charts/'.$oShape->getIndexedFilename())); //StrikeThrough XAxis
        $this->assertEquals(Font::UNDERLINE_NONE, $oXMLDoc->getElementAttribute($element, 'u', 'ppt/charts/'.$oShape->getIndexedFilename())); //Underline XAxis
        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill/a:srgbClr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('00FF00', $oXMLDoc->getElementAttribute($element, 'val', 'ppt/charts/'.$oShape->getIndexedFilename())); //Color YAxis
        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('false', $oXMLDoc->getElementAttribute($element, 'i', 'ppt/charts/'.$oShape->getIndexedFilename())); //Italic YAxis
        $this->assertEquals(16*100, $oXMLDoc->getElementAttribute($element, 'sz', 'ppt/charts/'.$oShape->getIndexedFilename())); //Size YAxis
        $this->assertEquals('noStrike', $oXMLDoc->getElementAttribute($element, 'strike', 'ppt/charts/'.$oShape->getIndexedFilename())); //StrikeThrough YAxis
        $this->assertEquals(Font::UNDERLINE_DASH, $oXMLDoc->getElementAttribute($element, 'u', 'ppt/charts/'.$oShape->getIndexedFilename())); //Underline YAxis
    }

    public function testTypeArea()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oArea = new Area();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFill()->setStartColor(new Color('FFAABBCC'));
        $oArea->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oArea);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:areaChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:areaChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeBar()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar = new Bar();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser/c:dPt/c:spPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser/c:tx/c:v';
        $this->assertEquals($oSeries->getTitle(), $oXMLDoc->getElement($element, 'ppt/charts/'.$oShape->getIndexedFilename())->nodeValue);
    }

    public function testTypeBar3D()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dPt/c:spPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:tx/c:v';
        $this->assertEquals($oSeries->getTitle(), $oXMLDoc->getElement($element, 'ppt/charts/'.$oShape->getIndexedFilename())->nodeValue);
    }

    public function testTypeBar3DSubScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeBar3DSuperScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeBar3DBarDirection()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oBar3D->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
        $oSeries = new Series('Downloads', $this->seriesData);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:barDir';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals(Bar3D::DIRECTION_HORIZONTAL, $oXMLDoc->getElementAttribute($element, 'val', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeLine()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:v';
        $this->assertEquals($oSeries->getTitle(), $oXMLDoc->getElement($element, 'ppt/charts/'.$oShape->getIndexedFilename())->nodeValue);
    }

    public function testTypeLineSubScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeLineSuperScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypePie()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie = new Pie();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oPie->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser/c:dPt/c:spPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser/c:tx/c:v';
        $this->assertEquals($oSeries->getTitle(), $oXMLDoc->getElement($element, 'ppt/charts/'.$oShape->getIndexedFilename())->nodeValue);
    }
    public function testTypePie3D()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dPt/c:spPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:tx/c:v';
        $this->assertEquals($oSeries->getTitle(), $oXMLDoc->getElement($element, 'ppt/charts/'.$oShape->getIndexedFilename())->nodeValue);
    }

    public function testTypePie3DExplosion()
    {
        $value = rand(1, 100);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oPie3D->setExplosion($value);
        $oSeries = new Series('Downloads', $this->seriesData);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:explosion';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals($value, $oXMLDoc->getElementAttribute($element, 'val', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypePie3DSubScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypePie3DSuperScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeScatter()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:tx/c:v';
        $this->assertEquals($oSeries->getTitle(), $oXMLDoc->getElement($element, 'ppt/charts/'.$oShape->getIndexedFilename())->nodeValue);
    }

    public function testTypeScatterSubScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeScatterSuperScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $oXMLDoc = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
}
