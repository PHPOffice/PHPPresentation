<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object'.EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPresentation Team')
                                  ->setTitle('Sample 06 Title')
                                  ->setSubject('Sample 06 Subject')
                                  ->setDescription('Sample 06 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPresentation->removeSlideByIndex(0);

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation);

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
$cell->createTextRun('Title row')->getFont()->setBold(true)->setSize(16);
$cell->getBorders()->getBottom()->setLineWidth(4)
                                ->setLineStyle(Border::LINE_SINGLE)
                                ->setDashStyle(Border::DASH_DASH);
$cell->getActiveParagraph()->getAlignment()
	->setMarginLeft(10);

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->setHeight(20);
$row->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)
               ->setRotation(90)
               ->setStartColor(new Color('FFE06B20'))
               ->setEndColor(new Color('FFFFFFFF'));
$oCell = $row->nextCell();
$oCell->createTextRun('R1C1')->getFont()->setBold(true);
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(20);
$oCell = $row->nextCell();
$oCell->createTextRun('R1C2')->getFont()->setBold(true);
$oCell = $row->nextCell();
$oCell->createTextRun('R1C3')->getFont()->setBold(true);

foreach ($row->getCells() as $cell) {
    $cell->getBorders()->getTop()->setLineWidth(4)
                                 ->setLineStyle(Border::LINE_SINGLE)
                                 ->setDashStyle(Border::DASH_DASH);
}

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('FFE06B20'))
               ->setEndColor(new Color('FFE06B20'));
$oCell = $row->nextCell();
$oCell->createTextRun('R2C1');
$oCell->getActiveParagraph()->getAlignment()
	->setMarginLeft(30)
	->setTextDirection(\PhpOffice\PhpPresentation\Style\Alignment::TEXT_DIRECTION_VERTICAL_270);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C2');
$oCell->getActiveParagraph()->getAlignment()
	->setMarginBottom(10)
	->setMarginTop(20)
	->setMarginRight(30)
	->setMarginLeft(40);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('FFE06B20'))
               ->setEndColor(new Color('FFE06B20'));
$oCell = $row->nextCell();
$oCell->createTextRun('R3C1');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(40);
$oCell = $row->nextCell();
$oCell->createTextRun('R3C2');
$oCell = $row->nextCell();
$oCell->createTextRun('R3C3');

// Add row
echo date('H:i:s') . ' Add row'.EOL;
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('FFE06B20'))
               ->setEndColor(new Color('FFE06B20'));
$cellC1 = $row->nextCell();
$textRunC1 = $cellC1->createTextRun('Link');
$textRunC1->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');
$cellC1->getActiveParagraph()->getAlignment()->setMarginLeft(50);
$cellC2 = $row->nextCell();
$textRunC2 = $cellC2->createTextRun('RichText with');
$textRunC2->getFont()->setBold(true);
$textRunC2->getFont()->setSize(12);
$textRunC2->getFont()->setColor(new Color('FF000000'));
$cellC2->createBreak();
$textRunC2 = $cellC2->createTextRun('Multiline');
$textRunC2->getFont()->setBold(true);
$textRunC2->getFont()->setSize(14);
$textRunC2->getFont()->setColor(new Color('FF0088FF'));
$cellC3 = $row->nextCell();
$textRunC3 = $cellC3->createTextRun('Link Github');
$textRunC3->getHyperlink()->setUrl('https://github.com')->setTooltip('GitHub');
$cellC3->createBreak();
$textRunC3 = $cellC3->createTextRun('Link Google');
$textRunC3->getHyperlink()->setUrl('https://google.com')->setTooltip('Google');

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}