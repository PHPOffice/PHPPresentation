<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Color;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPowerPoint Team')
                                  ->setTitle('Sample 05 Title')
                                  ->setSubject('Sample 05 Subject')
                                  ->setDescription('Sample 05 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

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
                   ->setColor( new Color( 'FFE06B20' ) );
$shape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPowerPoint/')
                      ->setTooltip('PHPPowerPoint');

// Create a shape (line)
$shape = $currentSlide->createLineShape(170, 180, 770, 180);
$shape->getBorder()->getColor()->setARGB( 'FFE06B20' );

// Create a shape (line)
$shape = $currentSlide->createLineShape(170, 580, 770, 580);
$shape->getBorder()->getColor()->setARGB( 'FFE06B20' );

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}