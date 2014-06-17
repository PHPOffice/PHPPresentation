<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPowerPoint Team')
                                  ->setTitle('Sample 04 Title')
                                  ->setSubject('Sample 04 Subject')
                                  ->setDescription('Sample 04 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Generate an image
echo date('H:i:s') . ' Generate an image'.EOL;
$gdImage = @imagecreatetruecolor(140, 20) or die('Cannot Initialize new GD image stream');
$textColor = imagecolorallocate($gdImage, 255, 255, 255);
imagestring($gdImage, 1, 5, 5,  'Created with PHPPowerPoint', $textColor);

// Add a drawing to the worksheet
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

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
