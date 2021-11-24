<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();
// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
    ->setLastModifiedBy('PHPPresentation Team')
    ->setTitle('Sample 21 AutoShape')
    ->setSubject('Sample 21 Subject')
    ->setDescription('Sample 21 Description')
    ->setKeywords('office 2007 openxml libreoffice odt php')
    ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

$autoShape = new AutoShape();
$autoShape->setType(AutoShape::TYPE_PENTAGON)
    ->setText('Step 1')
    ->setOffsetX(93)
    ->setOffsetY(30)
    ->setWidthAndHeight(175, 100);
$autoShape->getOutline()
    ->setWidth(0)
    ->getFill()
    ->setFillType(Fill::FILL_SOLID)
    ->setStartColor(new Color(Color::COLOR_BLACK));
$autoShape->getFill()
    ->setFillType(Fill::FILL_SOLID)
    ->setStartColor(new Color('804F81BD'));
$currentSlide->addShape($autoShape);

for ($inc = 1; $inc < 5; ++$inc) {
    $autoShape = new AutoShape();
    $autoShape->setType(AutoShape::TYPE_CHEVRON)
        ->setText('Step ' . ($inc + 1))
        ->setOffsetX(93 + $inc * 100)
        ->setOffsetY(30)
        ->setWidthAndHeight(175, 100);
    $autoShape->getOutline()
        ->setWidth($inc)
        ->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->setStartColor(new Color(Color::COLOR_BLACK));
    $autoShape->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->setStartColor(new Color('FF4F81BD'));
    $currentSlide->addShape($autoShape);
}

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
