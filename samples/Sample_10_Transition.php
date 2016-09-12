<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Slide\Transition;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPresentation Team')
                                  ->setTitle('Sample 10 Title')
                                  ->setSubject('Sample 10 Subject')
                                  ->setDescription('Sample 10 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$slide0 = $objPHPPresentation->getActiveSlide();

// Create a shape (drawing)
echo date('H:i:s') . ' Create a shape (drawing)'.EOL;
$shapeDrawing = $slide0->createDrawingShape();
$shapeDrawing->setName('PHPPresentation logo')
      ->setDescription('PHPPresentation logo')
      ->setPath('./resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shapeDrawing->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);
$shapeDrawing->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shapeRichText = $slide0->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$shapeRichText->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $shapeRichText->createTextRun('Thank you for using PHPPresentation!');
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
$slide1 = $objPHPPresentation->createSlide();
$slide1->addShape(clone $shapeDrawing);
$slide1->addShape(clone $shapeRichText);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
