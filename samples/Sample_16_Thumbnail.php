<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set Thumbnail
$objPHPPresentation->getPresentationProperties()->setThumbnailPath(__DIR__ . '\resources\phppowerpoint_logo.gif');

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$oSlide1 = $objPHPPresentation->getActiveSlide();
$oSlide1->addShape(clone $oShapeDrawing);
$oSlide1->addShape(clone $oShapeRichText);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
