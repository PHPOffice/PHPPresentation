<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);

/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . '../Classes/');

/** PHPPowerPoint */
include 'PHPPowerPoint.php';

/** PHPExcel */
include 'PHPExcel.php';

// Create new PHPPowerPoint object
echo date('H:i:s') . " Create new PHPPowerPoint object\n";
$objPHPPowerPoint = new PHPPowerPoint();

// Set properties
echo date('H:i:s') . " Set properties\n";
$objPHPPowerPoint->getProperties()->setCreator("Maarten Balliauw")
                                  ->setLastModifiedBy("Maarten Balliauw")
                                  ->setTitle("Office 2007 PPTX Test Document")
                                  ->setSubject("Office 2007 PPTX Test Document")
                                  ->setDescription("Test document for Office 2007 PPTX, generated using PHP classes.")
                                  ->setKeywords("office 2007 openxml php")
                                  ->setCategory("Test result file");

// Remove first slide
echo date('H:i:s') . " Remove first slide\n";
$objPHPPowerPoint->removeSlideByIndex(0);

// Create templated slide
echo date('H:i:s') . " Create templated slide\n";
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function

// Generate sample data for first chart
echo date('H:i:s') . " Generate sample data for first chart\n";
$series1Data = array('Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293);
$series2Data = array('Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379);

// Create a bar chart (that should be inserted in a shape)
echo date('H:i:s') . " Create a bar chart (that should be inserted in a chart shape)\n";
$bar3DChart = new PHPPowerPoint_Shape_Chart_Type_Bar3D();
$bar3DChart->addSeries( new PHPPowerPoint_Shape_Chart_Series('2009', $series1Data) );
$bar3DChart->addSeries( new PHPPowerPoint_Shape_Chart_Series('2010', $series2Data) );

// Create a shape (chart)
echo date('H:i:s') . " Create a shape (chart)\n";
$shape = $currentSlide->createChartShape();
$shape->setName('PHPPowerPoint Monthly Downloads')
      ->setResizeProportional(false)
      ->setHeight(550)
      ->setWidth(700)
      ->setOffsetX(120)
      ->setOffsetY(80)
      ->setIncludeSpreadsheet(true);
$shape->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);
$shape->getFill()->setFillType(PHPPowerPoint_Style_Fill::FILL_GRADIENT_LINEAR)
                 ->setStartColor(new PHPPowerPoint_Style_Color('FFCCCCCC'))
                 ->setEndColor(new PHPPowerPoint_Style_Color('FFFFFFFF'))
                 ->setRotation(270);
$shape->getBorder()->setLineStyle(PHPPowerPoint_Style_Border::LINE_SINGLE);
$shape->getTitle()->setText('PHPPowerPoint Monthly Downloads');
$shape->getTitle()->getFont()->setItalic(true);
$shape->getPlotArea()->getAxisX()->setTitle('Month');
$shape->getPlotArea()->getAxisY()->setTitle('Downloads');
$shape->getPlotArea()->setType($bar3DChart);
$shape->getView3D()->setRightAngleAxes(true);
$shape->getView3D()->setRotationX(20);
$shape->getView3D()->setRotationY(20);
$shape->getLegend()->getBorder()->setLineStyle(PHPPowerPoint_Style_Border::LINE_SINGLE);
$shape->getLegend()->getFont()->setItalic(true);

// Create templated slide
echo date('H:i:s') . " Create templated slide\n";
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function

// Generate sample data for second chart
echo date('H:i:s') . " Generate sample data for second chart\n";
$seriesData = array('Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7);

// Create a pie chart (that should be inserted in a shape)
echo date('H:i:s') . " Create a pie chart (that should be inserted in a chart shape)\n";
$pie3DChart = new PHPPowerPoint_Shape_Chart_Type_Pie3D();
$pie3DChart->addSeries( new PHPPowerPoint_Shape_Chart_Series('Downloads', $seriesData) );

// Create a shape (chart)
echo date('H:i:s') . " Create a shape (chart)\n";
$shape = $currentSlide->createChartShape();
$shape->setName('PHPPowerPoint Daily Downloads')
      ->setResizeProportional(false)
      ->setHeight(550)
      ->setWidth(700)
      ->setOffsetX(120)
      ->setOffsetY(80)
      ->setIncludeSpreadsheet(true);
$shape->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);
$shape->getFill()->setFillType(PHPPowerPoint_Style_Fill::FILL_GRADIENT_LINEAR)
                 ->setStartColor(new PHPPowerPoint_Style_Color('FFCCCCCC'))
                 ->setEndColor(new PHPPowerPoint_Style_Color('FFFFFFFF'))
                 ->setRotation(270);
$shape->getBorder()->setLineStyle(PHPPowerPoint_Style_Border::LINE_SINGLE);
$shape->getTitle()->setText('PHPPowerPoint Daily Downloads');
$shape->getTitle()->getFont()->setItalic(true);
$shape->getPlotArea()->setType($pie3DChart);
$shape->getView3D()->setRotationX(30);
$shape->getView3D()->setPerspective(30);
$shape->getLegend()->getBorder()->setLineStyle(PHPPowerPoint_Style_Border::LINE_SINGLE);
$shape->getLegend()->getFont()->setItalic(true);

// Save PowerPoint 2007 file
echo date('H:i:s') . " Write to PowerPoint2007 format\n";
$objWriter = PHPPowerPoint_IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$objWriter->save(str_replace('.php', '.pptx', __FILE__));

// Echo memory peak usage
echo date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB\r\n";

// Echo done
echo date('H:i:s') . " Done writing file.\r\n";



/**
 * Creates a templated slide
 *
 * @param PHPPowerPoint $objPHPPowerPoint
 * @return PHPPowerPoint_Slide
 */
function createTemplatedSlide(PHPPowerPoint $objPHPPowerPoint)
{
    // Create slide
    $slide = $objPHPPowerPoint->createSlide();

    // Add logo
    $slide->createDrawingShape()
          ->setName('PHPPowerPoint logo')
          ->setDescription('PHPPowerPoint logo')
          ->setPath('./resources/phppowerpoint_logo.gif')
          ->setHeight(40)
          ->setOffsetX(10)
          ->setOffsetY(720 - 10 - 40);

    // Return slide
    return $slide;
}
