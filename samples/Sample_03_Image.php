<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object'.EOL;
$objPHPPresentation = new PhpPresentation();

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

// Generate an image
echo date('H:i:s') . ' Generate an image'.EOL;
$gdImage = @imagecreatetruecolor(140, 20) or die('Cannot Initialize new GD image stream');
$textColor = imagecolorallocate($gdImage, 255, 255, 255);
imagestring($gdImage, 1, 5, 5,  'Created with PHPPresentation', $textColor);

// Add a generated drawing to the slide
echo date('H:i:s') . ' Add a drawing to the slide'.EOL;
$shape = new MemoryDrawing();
$shape->setName('Sample image')
      ->setDescription('Sample image')
      ->setImageResource($gdImage)
      ->setRenderingFunction(MemoryDrawing::RENDERING_JPEG)
      ->setMimeType(MemoryDrawing::MIMETYPE_DEFAULT)
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$currentSlide->addShape($shape);

// Add a file drawing (GIF) to the slide
$shape = new Drawing();
$shape->setName('PHPPresentation logo')
      ->setDescription('PHPPresentation logo')
      ->setPath('./resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(100);
$currentSlide->addShape($shape);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
