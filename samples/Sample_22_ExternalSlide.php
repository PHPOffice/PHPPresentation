<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();
$objPHPPresentation->removeSlideByIndex(0);

$oReader = PhpOffice\PhpPresentation\IOFactory::createReader('PowerPoint2007');
$oPresentation04 = $oReader->load(__DIR__ . '/results/Sample_04_Table.pptx');

foreach ($oPresentation04->getAllSlides() as $oSlide) {
    $objPHPPresentation->addExternalSlide($oSlide);
}

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
