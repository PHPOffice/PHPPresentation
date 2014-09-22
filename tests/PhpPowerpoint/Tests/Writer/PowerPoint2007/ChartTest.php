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

namespace PhpOffice\PhpPowerpoint\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\Chart\Series;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Line;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Tests\TestHelperDOCX;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Chart;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class ChartTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Invalid \PhpOffice\PhpPowerpoint\Shape\Chart object passed.
     */
    public function testConstructException()
    {
        $oChart = new Chart();
        $oChart->writeChart();
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage The chart type provided could not be rendered.
     */
    public function testPlotAreaBadType()
    {
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $stub = $this->getMockForAbstractClass('PhpOffice\PhpPowerpoint\Shape\Chart\Type\AbstractType');
        $oShape->getPlotArea()->setType($stub);
        
        TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    }
    
    public function testTypeBar3D()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
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
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
    
    public function testTypeBar3DSuperScript()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
    
    public function testTypeLine()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $seriesData);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
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
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
    
    public function testTypeLineSuperScript()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
    
    public function testTypePie3D()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
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
    
    public function testTypePie3DSubScript()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
    
    public function testTypePie3DSuperScript()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }

    public function testTypeScatter()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $seriesData);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
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
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
    
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
    
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('-25000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
    
    public function testTypeScatterSuperScript()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );
        
        $oPHPPowerPoint = new PhpPowerpoint();
        $oSlide = $oPHPPowerPoint->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);
        
        $oXMLDoc = TestHelperDOCX::getDocument($oPHPPowerPoint, 'PowerPoint2007');
        
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $this->assertEquals('30000', $oXMLDoc->getElementAttribute($element, 'baseline', 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
}
