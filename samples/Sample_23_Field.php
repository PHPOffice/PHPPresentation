<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\RichText\Field;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')->setLastModifiedBy('PHPPresentation Team')->setTitle('Sample 23 Title')->setSubject('Sample 23 Subject')->setDescription('Sample 23 Description')->setKeywords('office 2007 openxml libreoffice odt php')->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

// Create a paragraph part written and part computed
echo date('H:i:s') . ' Create a shape (rich text) holding fields' . EOL;
$shape = $currentSlide->createRichTextShape()
    ->setHeight(100)
    ->setWidth(600)
    ->setOffsetX(170)
    ->setOffsetY(180);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$paragraph = $shape->getActiveParagraph();
$paragraph->createTextRun('page ');
$paragraph->createField(Field::TYPE_SLIDENUM, '<nr.>')->getFont()->setBold(true)->setColor(new Color('FFE06B20'));
$paragraph->createTextRun(' of ');
$paragraph->createField(Field::TYPE_SLIDECOUNT, '2')->getFont()->setBold(true)->setColor(new Color('FFE06B20'));

// A field can stand on its own as well as inside a sentence
$paragraph = $shape->createParagraph();
$paragraph->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$paragraph->createField(Field::TYPE_DATETIME, '01-01-25');

// Create a second slide, so that the count is a number worth computing
echo date('H:i:s') . ' Create slide' . EOL;
$objPHPPresentation->createSlide()->createRichTextShape()
    ->setHeight(100)
    ->setWidth(600)
    ->setOffsetX(170)
    ->setOffsetY(180)
    ->createTextRun('The second slide');

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
