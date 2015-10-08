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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
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
class ChartAreaTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }
    
    public function testTypeArea()
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
        $oArea = new Area();
        $oSeries = new Series('Downloads', $seriesData);
        $oSeries->getFill()->setStartColor(new Color('FFAABBCC'));
        $oArea->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oArea);
    
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/slides/slide1.xml'));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:areaChart';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:areaChart/c:ser';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/charts/'.$oShape->getIndexedFilename()));
    }
}
