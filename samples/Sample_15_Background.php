<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide\Background\Color;
use PhpOffice\PhpPresentation\Style\Color as StyleColor;
use \PhpOffice\PhpPresentation\Slide\Background\Image;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$oSlide1 = $objPHPPresentation->getActiveSlide();
$oSlide1->addShape(clone $oShapeDrawing);
$oSlide1->addShape(clone $oShapeRichText);

// Slide > Background > Color
$oBkgColor = new Color();
$oBkgColor->setColor(new StyleColor(StyleColor::COLOR_DARKGREEN));
$oSlide1->setBackground($oBkgColor);

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$oSlide2 = $objPHPPresentation->createSlide();
$oSlide2->addShape(clone $oShapeDrawing);
$oSlide2->addShape(clone $oShapeRichText);

// Slide > Background > Image
/*
 * @link : http://publicdomainarchive.com/public-domain-images-cave-red-rocks-light-beam-cavern/
 */
$oBkgImage = new Image();
$oBkgImage->setPath('./resources/background.jpg');
$oSlide2->setBackground($oBkgImage);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
