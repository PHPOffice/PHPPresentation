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

namespace PhpOffice\PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\Chart;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
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
    
    public function testTypeBar()
    {
        $seriesData = array(
            'A' => 1,
            'B' => 2,
            'C' => 4,
            'D' => 3,
            'E' => 2,
        );

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar = new Bar();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');

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
}
