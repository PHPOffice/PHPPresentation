<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
    ->setLastModifiedBy('PHPPresentation Team')
    ->setTitle('Sample 01 Title')
    ->setSubject('Sample 01 Subject')
    ->setDescription('Sample 01 Description')
    ->setKeywords('office 2007 openxml libreoffice odt php')
    ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

// Create a shape (drawing)
echo date('H:i:s') . ' Create a shape (drawing)' . EOL;
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPresentation logo')
    ->setDescription('PHPPresentation logo')
    ->setPath(__DIR__ . '/resources/phppowerpoint_logo.gif')
    ->setHeight(36)
    ->setOffsetX(10)
    ->setOffsetY(10);
$shape->getShadow()->setVisible(true)
    ->setDirection(45)
    ->setDistance(10);
$shape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)' . EOL;
$shape = $currentSlide->createRichTextShape()
    ->setHeight(300)
    ->setWidth(600)
    ->setOffsetX(170)
    ->setOffsetY(100);
$shape->getActiveParagraph()->getAlignment()
    ->setHorizontal(Alignment::HORIZONTAL_CENTER);
$textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
    ->setSize(60)
    ->setColor(new Color('FFE06B20'));

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)' . EOL;
$shape = $currentSlide->createRichTextShape()
    ->setHeight(300)
    ->setWidth(600)
    ->setOffsetX(170)
    ->setOffsetY(550);
$shape->getActiveParagraph()->getAlignment()
    ->setHorizontal(Alignment::HORIZONTAL_RIGHT)
    ->setIsRTL(true);
$textRun = $shape->createTextRun('تست فونت فارسی');
$textRun->getFont()
    ->setBold(true)
    ->setSize(60)
    ->setColor(new Color('FFE06B20'))
    ->setFormat(Font::FORMAT_COMPLEX_SCRIPT)
    ->setName('B Nazanin');

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
