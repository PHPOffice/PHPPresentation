<?php
include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();
// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
    ->setLastModifiedBy('PHPPresentation Team')
    ->setTitle('Sample 21 SlideMaster')
    ->setSubject('Sample 21 Subject')
    ->setDescription('Sample 21 Description')
    ->setKeywords('office 2007 openxml libreoffice odt php')
    ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

//
$oAutoShape = new AutoShape();
$oAutoShape->setType(AutoShape::TYPE_PENTAGON)
    ->setText('Step 1')
    ->setOffsetX(93)
    ->setOffsetY(30)
    ->setWidthAndHeight(175, 100);
$currentSlide->addShape($oAutoShape);

for ($inc = 1 ; $inc <5 ; $inc++) {
    $oAutoShape = new AutoShape();
    $oAutoShape->setType(AutoShape::TYPE_CHEVRON)
        ->setText('Step '. ($inc + 1))
        ->setOffsetX(93 + $inc * 50)
        ->setOffsetY(30)
        ->setWidthAndHeight(175, 100);
    $currentSlide->addShape($oAutoShape);
}

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}