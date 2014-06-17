<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\IOFactory;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPowerPoint Team')
                                  ->setTitle('Sample 03 Title')
                                  ->setSubject('Sample 03 Subject')
                                  ->setDescription('Sample 03 Description')
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
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $shape->createTextRun('Thank you for using PHPPowerPoint!');
$textRun->getFont()->setBold(true)
                   ->setSize(60)
                   ->setColor( new Color( 'FFE06B20' ) );

// Save serialized file
$basename = basename(__FILE__, '.php');
echo date('H:i:s') . ' Write to serialized format'.EOL;
$objWriter = IOFactory::createWriter($objPHPPowerPoint, 'Serialized');
$objWriter->save('results/'.basename(__FILE__, '.php').'.phppt');

// Read from serialized file
echo date('H:i:s') . ' Read from serialized format'.EOL;
$objPHPPowerPointLoaded = IOFactory::load('results/'.basename(__FILE__, '.php').'.phppt');

// Save file
echo write($objPHPPowerPointLoaded, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
