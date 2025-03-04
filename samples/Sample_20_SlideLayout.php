<?php

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Style\Color;

include_once 'Sample_Header.php';

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
    ->setLastModifiedBy('PHPPresentation Team')
    ->setTitle('Sample 20 SlideLayout')
    ->setSubject('Sample 20 Subject')
    ->setDescription('Sample 20 Description')
    ->setKeywords('office 2007 openxml libreoffice odt php')
    ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

echo date('H:i:s') . ' Create SlideLayout' . EOL;
$slideLayout = $objPHPPresentation->getAllMasterSlides()[0]->createSlideLayout();
$slideLayout->setLayoutName('Sample Layout');

echo date('H:i:s') . ' Create Footer' . EOL;
$footerTextShape = $slideLayout->createRichTextShape();
$footerTextShape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_FOOTER));

$footerTextShape
    ->setOffsetX(77)
    ->setOffsetY(677)
    ->setWidth(448)
    ->setHeight(23);

$footerTextRun = $footerTextShape->createTextRun('Footer placeholder');
$footerTextRun->getFont()
    ->setName('Calibri')
    ->setSize(9)
    ->setColor(new Color(Color::COLOR_DARKGREEN))
    ->setBold(true);

echo date('H:i:s') . ' Create SlideNumber' . EOL;

$numberTextShape = $slideLayout->createRichTextShape();
$numberTextShape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_SLIDENUM));

$numberTextShape
    ->setOffsetX(43)
    ->setOffsetY(677)
    ->setWidth(43)
    ->setHeight(23);

$numberTextRun = $numberTextShape->createTextRun('');
$numberTextRun->getFont()
    ->setName('Calibri')
    ->setSize(9)
    ->setColor(new Color(Color::COLOR_DARKGREEN))
    ->setBold(true);

echo date('H:i:s') . ' Apply Layout' . EOL;
$currentSlide->setSlideLayout($slideLayout);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'));
if (!CLI) {
    include_once 'Sample_Footer.php';
}
