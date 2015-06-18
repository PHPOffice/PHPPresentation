<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Slide\Transition;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object' . EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPowerPoint Team')
                                  ->setTitle('Sample 10 Title')
                                  ->setSubject('Sample 10 Subject')
                                  ->setDescription('Sample 10 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$slide0 = $objPHPPowerPoint->getActiveSlide();

// Create a shape (drawing)
echo date('H:i:s') . ' Create a shape (drawing)'.EOL;
$shapeDrawing = $slide0->createDrawingShape();
$shapeDrawing->setName('PHPPowerPoint logo')
      ->setDescription('PHPPowerPoint logo')
      ->setPath('./resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shapeDrawing->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);
$shapeDrawing->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPowerPoint/')->setTooltip('PHPPowerPoint');

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shapeRichText = $slide0->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$shapeRichText->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $shapeRichText->createTextRun('Thank you for using PHPPowerPoint!');
$textRun->getFont()->setBold(true)
                   ->setSize(60)
                   ->setColor( new Color( 'FFE06B20' ) );

$oTransition = new Transition();
$oTransition->setManualTrigger(false);
$oTransition->setTimeTrigger(true, 4000);
$oTransition->setTransitionType(Transition::TRANSITION_SPLIT_IN_VERTICAL);
$slide0->setTransition($oTransition);

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$slide1 = $objPHPPowerPoint->createSlide();
$slide1->addShape(clone $shapeDrawing);
$slide1->addShape(clone $shapeRichText);

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
