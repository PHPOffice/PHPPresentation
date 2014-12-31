<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Line;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPowerpoint\Shape\Chart\Series;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Style\Shadow;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPowerPoint Team')
                                  ->setTitle('Sample 07 Title')
                                  ->setSubject('Sample 07 Subject')
                                  ->setDescription('Sample 07 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPowerPoint->removeSlideByIndex(0);

// Set Style
$oFill = new Fill();
$oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

$oShadow = new Shadow();
$oShadow->setVisible(true)->setDirection(45)->setDistance(10);

// Generate sample data for chart
echo date('H:i:s') . ' Generate sample data for chart'.EOL;
$seriesData = array('Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7);

// Create templated slide
echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint);

// Create a line chart (that should be inserted in a shape)
echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)'.EOL;
$lineChart = new Line();
$series = new Series('Downloads', $seriesData);
$series->setShowSeriesName(true);
$series->setShowValue(true);
$lineChart->addSeries($series);

// Create a shape (chart)
echo date('H:i:s') . ' Create a shape (chart)'.EOL;
$shape = $currentSlide->createChartShape();
$shape->setName('PHPPowerPoint Daily Downloads')
      ->setResizeProportional(false)
      ->setHeight(550)
      ->setWidth(700)
      ->setOffsetX(120)
      ->setOffsetY(80);
$shape->setShadow($oShadow);
$shape->setFill($oFill);
$shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getTitle()->setText('PHPPowerPoint Daily Downloads');
$shape->getTitle()->getFont()->setItalic(true);
$shape->getPlotArea()->setType($lineChart);
$shape->getView3D()->setRotationX(30);
$shape->getView3D()->setPerspective(30);
$shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getLegend()->getFont()->setItalic(true);

// Create templated slide
echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint);

// Create a line chart (that should be inserted in a shape)
echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)'.EOL;
$lineChart1 = clone $lineChart;

// Create a shape (chart)
echo date('H:i:s') . ' Create a shape (chart)'.EOL;
echo date('H:i:s') . ' Differences with previous : Values on right axis and Legend hidden'.EOL;
$shape1 = clone $shape;
$shape1->getLegend()->setVisible(false);
$shape1->setName('PHPPowerPoint Weekly Downloads');
$shape1->getTitle()->setText('PHPPowerPoint Weekly Downloads');
$shape1->getPlotArea()->setType($lineChart1);
$shape1->getPlotArea()->getAxisY()->setFormatCode('#,##0');
$currentSlide->addShape($shape1);

// Save file
echo EOL.write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);

if (!CLI) {
    include_once 'Sample_Footer.php';
}
