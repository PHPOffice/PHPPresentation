<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Media;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

// Add a video to the slide
$shape = new Media();
$shape->setName('Video')
    ->setDescription('Video')
    ->setPath(
        __DIR__ . '/resources/sintel_trailer-480p' .
        ('WIN' === strtoupper(substr(PHP_OS, 0, 3)) ? '.wmv' : '.ogv')
    )
    ->setResizeProportional(false)
    ->setHeight(90)
    ->setWidth(90)
    ->setOffsetX(10)
    ->setOffsetY(300);
$currentSlide->addShape($shape);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
