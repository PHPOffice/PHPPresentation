<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

// Generate an image
echo date('H:i:s') . ' Generate an image' . EOL;
$gdImage = @imagecreatetruecolor(140, 20) or exit('Cannot Initialize new GD image stream');
$textColor = imagecolorallocate($gdImage, 255, 255, 255);
imagestring($gdImage, 1, 5, 5, 'Created with PHPPresentation', $textColor);

// Add a generated drawing to the slide
echo date('H:i:s') . ' Add a drawing to the slide' . EOL;
$shape = new Drawing\Gd();
$shape->setName('Image GD')
    ->setDescription('Image GD')
    ->setImageResource($gdImage)
    ->setMimeType(Drawing\Gd::MIMETYPE_DEFAULT)
    ->setHeight(36)
    ->setOffsetX(10)
    ->setOffsetY(10);
$currentSlide->addShape($shape);

// Add a file drawing (GIF) to the slide
$shape = new Drawing\File();
$shape->setName('Image File')
    ->setDescription('Image File')
    ->setPath(__DIR__ . '/resources/phppowerpoint_logo.gif')
    ->setHeight(36)
    ->setOffsetX(10)
    ->setOffsetY(100);
$currentSlide->addShape($shape);

// Add a file drawing (Zip) to the slide
if (file_exists(__DIR__ . '/resources/Sample_12.pptx')) {
    $shape = new Drawing\ZipFile();
    $shape->setName('Image ZipFile')
        ->setDescription('Image ZipFile')
        ->setPath('zip://' . __DIR__ . '/resources/Sample_12.pptx#ppt/media/phppowerpoint_logo1.gif')
        ->setResizeProportional(false)
        ->setHeight(36)
        ->setWidth(36)
        ->setOffsetX(10)
        ->setOffsetY(150);
    $currentSlide->addShape($shape);
}

// Add a file drawing (JPEG) to the slide
$shape = new Drawing\Base64();
$shape->setName('Image Base64')
    ->setDescription('Image Base64')
    ->setData(file_get_contents(__DIR__ . '/resources/base64.txt'))
    ->setResizeProportional(false)
    ->setHeight(36)
    ->setWidth(36)
    ->setOffsetX(10)
    ->setOffsetY(200);
$currentSlide->addShape($shape);

// Add a file drawing (PNG transparent) to the slide
$fill = new Fill();
$fill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));

$shape = new Drawing\File();
$shape->setName('Image File PNG')
    ->setDescription('Image File PNG')
    ->setPath(__DIR__ . '/resources/logo_ubuntu_transparent.png')
    ->setHeight(100)
    ->setOffsetX(10)
    ->setOffsetY(250)
    ->setFill($fill);
$currentSlide->addShape($shape);

// Add a file drawing (SVG) to the slide
$shape = new Drawing\File();
$shape->setName('Image File SVG')
    ->setDescription('Image File SVG')
    ->setPath(__DIR__ . '/resources/tiger.svg')
    ->setHeight(100)
    ->setWidth(100)
    ->setOffsetX(10)
    ->setOffsetY(360);
$currentSlide->addShape($shape);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
