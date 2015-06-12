<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\Drawing;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Generate an image
echo date('H:i:s') . ' Generate an image'.EOL;
$gdImage = @imagecreatetruecolor(140, 20) or die('Cannot Initialize new GD image stream');
$textColor = imagecolorallocate($gdImage, 255, 255, 255);
imagestring($gdImage, 1, 5, 5,  'Created with PHPPowerPoint', $textColor);

// Add a generated drawing to the slide
echo date('H:i:s') . ' Add a drawing to the worksheet'.EOL;
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
$shape->setName('PHPPowerPoint logo')
      ->setDescription('PHPPowerPoint logo')
      ->setPath('./resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(100);
$currentSlide->addShape($shape);

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
