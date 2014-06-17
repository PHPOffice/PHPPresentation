<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Fill;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPowerPoint Team')
                                  ->setTitle('Sample 06 Title')
                                  ->setSubject('Sample 06 Subject')
                                  ->setDescription('Sample 06 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPowerPoint->removeSlideByIndex(0);

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint);

// Create a shape (table)
echo date('H:i:s') . ' Create a shape (table)'.EOL;
$shape = $currentSlide->createTableShape(3);
$shape->setHeight(200);
$shape->setWidth(600);
$shape->setOffsetX(150);
$shape->setOffsetY(300);

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)
               ->setRotation(90)
               ->setStartColor(new Color('FFE06B20'))
               ->setEndColor(new Color('FFFFFFFF'));
$cell = $row->nextCell();
$cell->setColSpan(3);
$cell->createTextRun('Title row')->getFont()->setBold(true)
                                            ->setSize(16);
$cell->getBorders()->getBottom()->setLineWidth(4)
                                ->setLineStyle(Border::LINE_SINGLE)
                                ->setDashStyle(Border::DASH_DASH);

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->setHeight(20);
$row->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)
               ->setRotation(90)
               ->setStartColor(new Color( 'FFE06B20' ))
               ->setEndColor(new Color('FFFFFFFF'));
$row->nextCell()->createTextRun('R1C1')->getFont()->setBold(true);
$row->nextCell()->createTextRun('R1C2')->getFont()->setBold(true);
$row->nextCell()->createTextRun('R1C3')->getFont()->setBold(true);

foreach ($row->getCells() as $cell) {
    $cell->getBorders()->getTop()->setLineWidth(4)
                                 ->setLineStyle(Border::LINE_SINGLE)
                                 ->setDashStyle(Border::DASH_DASH);
}

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color( 'FFE06B20' ))
               ->setEndColor(new Color('FFE06B20'));
$row->nextCell()->createTextRun('R2C1');
$row->nextCell()->createTextRun('R2C2');
$row->nextCell()->createTextRun('R2C3');

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color( 'FFE06B20' ))
               ->setEndColor(new Color('FFE06B20'));
$row->nextCell()->createTextRun('R3C1');
$row->nextCell()->createTextRun('R3C2');
$row->nextCell()->createTextRun('R3C3');

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}