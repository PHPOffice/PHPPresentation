<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide\Animation;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

$oDrawing1 = clone $oShapeDrawing;
$oRichText1 = clone $oShapeRichText;

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$oSlide1 = $objPHPPresentation->getActiveSlide();
$oSlide1->addShape($oDrawing1);
$oSlide1->addShape($oRichText1);

$oAnimation1 = new Animation();
$oAnimation1->addShape($oDrawing1);
$oSlide1->addAnimation($oAnimation1);

$oAnimation2 = new Animation();
$oAnimation2->addShape($oRichText1);
$oSlide1->addAnimation($oAnimation2);

$oDrawing2 = clone $oShapeDrawing;
$oRichText2 = clone $oShapeRichText;

$oSlide2 = $objPHPPresentation->createSlide();
$oSlide2->addShape($oDrawing2);
$oSlide2->addShape($oRichText2);

$oAnimation4 = new Animation();
$oAnimation4->addShape($oRichText2);
$oSlide2->addAnimation($oAnimation4);

$oAnimation3 = new Animation();
$oAnimation3->addShape($oDrawing2);
$oSlide2->addAnimation($oAnimation3);

$oDrawing3 = clone $oShapeDrawing;
$oRichText3 = clone $oShapeRichText;

$oSlide3 = $objPHPPresentation->createSlide();
$oSlide3->addShape($oDrawing3);
$oSlide3->addShape($oRichText3);

$oAnimation5 = new Animation();
$oAnimation5->addShape($oRichText3);
$oAnimation5->addShape($oDrawing3);
$oSlide3->addAnimation($oAnimation5);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
