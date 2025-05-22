<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set the zoom to 200%
$objPHPPresentation->getPresentationProperties()->setZoom(3);

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();
$currentSlide->addShape(clone $oShapeDrawing);
$currentSlide->addShape(clone $oShapeRichText);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
