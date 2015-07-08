<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getProperties()->setCreator('PHPOffice')->setLastModifiedBy('PHPPresentation Team')->setTitle('Sample 07 Title')->setSubject('Sample 07 Subject')->setDescription('Sample 07 Description')->setKeywords('office 2007 openxml libreoffice odt php')->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide' . EOL;
$objPHPPresentation->removeSlideByIndex(0);

// Set Style
$oFill = new Fill();
$oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

$oShadow = new Shadow();
$oShadow->setVisible(true)->setDirection(45)->setDistance(10);

// Generate sample data for chart
echo date('H:i:s') . ' Generate sample data for chart' . EOL;
$seriesData = array(
	'Monday' => 12,
	'Tuesday' => 15,
	'Wednesday' => 13,
	'Thursday' => 17,
	'Friday' => 14,
	'Saturday' => 9,
	'Sunday' => 7
);

// Create templated slide
echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
$lineChart = new Line();
$series = new Series('Downloads', $seriesData);
$series->setShowSeriesName(true);
$series->setShowValue(true);
$lineChart->addSeries($series);

// Create a shape (chart)
echo date('H:i:s') . ' Create a shape (chart)' . EOL;
$shape = $currentSlide->createChartShape();
$shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
$shape->setShadow($oShadow);
$shape->setFill($oFill);
$shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getTitle()->setText('PHPPresentation Daily Downloads');
$shape->getTitle()->getFont()->setItalic(true);
$shape->getPlotArea()->setType($lineChart);
$shape->getView3D()->setRotationX(30);
$shape->getView3D()->setPerspective(30);
$shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
$shape->getLegend()->getFont()->setItalic(true);

// Create templated slide
echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
$lineChart1 = clone $lineChart;

// Create a shape (chart)
echo date('H:i:s') . ' Create a shape (chart)' . EOL;
echo date('H:i:s') . ' Differences with previous : Values on right axis and Legend hidden' . EOL;
$shape1 = clone $shape;
$shape1->getLegend()->setVisible(false);
$shape1->setName('PHPPresentation Weekly Downloads');
$shape1->getTitle()->setText('PHPPresentation Weekly Downloads');
$shape1->getPlotArea()->setType($lineChart1);
$shape1->getPlotArea()->getAxisY()->setFormatCode('#,##0');
$currentSlide->addShape($shape1);

// Create templated slide
echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
$lineChart2 = clone $lineChart;
$series2 = $lineChart2->getData();
$series2[0]->getFill()->setFillType(Fill::FILL_SOLID);
$lineChart2->setData($series2);

// Create a shape (chart)
echo date('H:i:s') . ' Create a shape (chart)' . EOL;
echo date('H:i:s') . ' Differences with previous : Values on right axis and Legend hidden' . EOL;
$shape2 = clone $shape;
$shape2->getLegend()->setVisible(false);
$shape2->setName('PHPPresentation Weekly Downloads');
$shape2->getTitle()->setText('PHPPresentation Weekly Downloads');
$shape2->getPlotArea()->setType($lineChart2);
$shape2->getPlotArea()->getAxisY()->setFormatCode('#,##0');
$currentSlide->addShape($shape2);

// Save file
echo EOL . write($objPHPPresentation, basename(__FILE__, '.php'), $writers);

if (!CLI) {
	include_once 'Sample_Footer.php';
}
