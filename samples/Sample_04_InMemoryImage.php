<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;

// Create new PHPPowerPoint object
echo date('H:i:s') . " Create new PHPPowerPoint object\n";
$objPHPPowerPoint = new \PhpOffice\PhpPowerpoint\PhpPowerpoint();

// Set properties
echo date('H:i:s') . " Set properties\n";
$objPHPPowerPoint->getProperties()->setCreator("Maarten Balliauw")
                                  ->setLastModifiedBy("Maarten Balliauw")
                                  ->setTitle("Office 2007 PPTX Test Document")
                                  ->setSubject("Office 2007 PPTX Test Document")
                                  ->setDescription("Test document for Office 2007 PPTX, generated using PHP classes.")
                                  ->setKeywords("office 2007 openxml php")
                                  ->setCategory("Test result file");

// Create slide
echo date('H:i:s') . " Create slide\n";
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Generate an image
echo date('H:i:s') . " Generate an image\n";
$gdImage = @imagecreatetruecolor(140, 20) or die('Cannot Initialize new GD image stream');
$textColor = imagecolorallocate($gdImage, 255, 255, 255);
imagestring($gdImage, 1, 5, 5,  'Created with PHPPowerPoint', $textColor);

// Add a drawing to the worksheet
echo date('H:i:s') . " Add a drawing to the worksheet\n";
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
