<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Color;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new \PhpOffice\PhpPowerpoint\PHPPowerPoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('Maarten Balliauw')
                                  ->setLastModifiedBy('Maarten Balliauw')
                                  ->setTitle('Office 2007 PPTX Test Document')
                                  ->setSubject('Office 2007 PPTX Test Document')
                                  ->setDescription('Test document for Office 2007 PPTX, generated using PHP classes.')
                                  ->setKeywords('office 2007 openxml php')
                                  ->setCategory('Test result file');

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Create a shape (drawing)
echo date('H:i:s') . ' Create a shape (drawing)'.EOL;
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPowerPoint logo')
      ->setDescription('PHPPowerPoint logo')
      ->setPath('./resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shape->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape()
      ->setHeight(400)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180)
      ->setInsetTop(50)
      ->setInsetBottom(50);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $shape->createTextRun('Thank you for using PHPPowerPoint!');
$textRun->getFont()->setBold(true)
                   ->setSize(60)
                   ->setColor( new Color( 'FFC00000' ) );
$shape->getHyperlink()->setUrl('http://phppowerpoint.codeplex.com')
                      ->setTooltip('PHPPowerPoint');

// Create a shape (line)
$shape = $currentSlide->createLineShape(170, 180, 770, 180);
$shape->getBorder()->getColor()->setARGB('FFC00000');

// Create a shape (line)
$shape = $currentSlide->createLineShape(170, 580, 770, 580);
$shape->getBorder()->getColor()->setARGB('FFC00000');

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}