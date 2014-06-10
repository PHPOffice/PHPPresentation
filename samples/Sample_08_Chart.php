<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Series;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new \PhpOffice\PhpPowerpoint\PHPPowerPoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('Maarten Balliauw')
                                  ->setLastModifiedBy('Maarten Balliauw')
                                  ->setTitle('Office 2007 PPTX Test Document')
                                  ->setSubject('Office 2007 PPTX Test Document')
                                  ->setDescription("Test document for Office 2007 PPTX, generated using PHP classes.")
                                  ->setKeywords("office 2007 openxml php")
                                  ->setCategory("Test result file");

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPowerPoint->removeSlideByIndex(0);

// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function

// Generate sample data for first chart
echo date('H:i:s') . ' Generate sample data for first chart'.EOL;
$series1Data = array('Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293);
$series2Data = array('Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379);

// Create a bar chart (that should be inserted in a shape)
echo date('H:i:s') . ' Create a bar chart (that should be inserted in a chart shape)'.EOL;
$bar3DChart = new Bar3D();
$bar3DChart->addSeries( new Series('2009', $series1Data) );
$bar3DChart->addSeries( new Series('2010', $series2Data) );

// Create a shape (chart)
echo date('H:i:s') . ' Create a shape (chart)'.EOL;
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
$shape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)
                 ->setStartColor(new Color('FFCCCCCC'))
                 ->setEndColor(new Color('FFFFFFFF'))
                 ->setRotation(270);
$shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getTitle()->setText('PHPPowerPoint Monthly Downloads');
$shape->getTitle()->getFont()->setItalic(true);
$shape->getPlotArea()->getAxisX()->setTitle('Month');
$shape->getPlotArea()->getAxisY()->setTitle('Downloads');
$shape->getPlotArea()->setType($bar3DChart);
$shape->getView3D()->setRightAngleAxes(true);
$shape->getView3D()->setRotationX(20);
$shape->getView3D()->setRotationY(20);
$shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getLegend()->getFont()->setItalic(true);

// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function

// Generate sample data for second chart
echo date('H:i:s') . ' Generate sample data for second chart'.EOL;
$seriesData = array('Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7);

// Create a pie chart (that should be inserted in a shape)
echo date('H:i:s') . ' Create a pie chart (that should be inserted in a chart shape)'.EOL;
$pie3DChart = new Pie3D();
$pie3DChart->addSeries( new Series('Downloads', $seriesData) );

// Create a shape (chart)
echo date('H:i:s') . ' Create a shape (chart)'.EOL;
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
$shape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)
                 ->setStartColor(new Color('FFCCCCCC'))
                 ->setEndColor(new Color('FFFFFFFF'))
                 ->setRotation(270);
$shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getTitle()->setText('PHPPowerPoint Daily Downloads');
$shape->getTitle()->getFont()->setItalic(true);
$shape->getPlotArea()->setType($pie3DChart);
$shape->getView3D()->setRotationX(30);
$shape->getView3D()->setPerspective(30);
$shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getLegend()->getFont()->setItalic(true);

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}



/**
 * Creates a templated slide
 *
 * @param PHPPowerPoint $objPHPPowerPoint
 * @return PHPPowerPoint_Slide
 */
function createTemplatedSlide(PhpPowerpoint $objPHPPowerPoint)
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
