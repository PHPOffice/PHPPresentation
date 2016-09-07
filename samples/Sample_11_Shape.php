<?php

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;

include_once 'Sample_Header.php';

function fnSlideRichText(PhpPresentation $objPHPPresentation)
{
    // Create templated slide
    echo date('H:i:s') . ' Create templated slide' . EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text)' . EOL;
    $shape = $currentSlide->createRichTextShape();
    $shape->setHeight(100);
    $shape->setWidth(600);
    $shape->setOffsetX(100);
    $shape->setOffsetY(100);
    $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $textRun = $shape->createTextRun('RichText with');
    $textRun->getFont()->setBold(true);
    $textRun->getFont()->setSize(28);
    $textRun->getFont()->setColor(new Color('FF000000'));

    $shape->createBreak();

    $textRun = $shape->createTextRun('Multiline');
    $textRun->getFont()->setBold(true);
    $textRun->getFont()->setSize(60);
    $textRun->getFont()->setColor(new Color('FF000000'));
}

function fnSlideRichTextLineSpacing(PhpPresentation $objPHPPresentation)
{
    // Create templated slide
    echo date('H:i:s') . ' Create templated slide' . EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text) with line spacing (100)' . EOL;
    $shape = $currentSlide->createRichTextShape();
    $shape->setHeight(100);
    $shape->setWidth(400);
    $shape->setOffsetX(100);
    $shape->setOffsetY(100);
    $shape->getActiveParagraph()->setLineSpacing(100);

    $shape->createTextRun('RichText with');
    $shape->createBreak();
    $shape->createTextRun('Line Spacing 100');

    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text) with line spacing (200)' . EOL;
    $shape = $currentSlide->createRichTextShape();
    $shape->setHeight(100);
    $shape->setWidth(400);
    $shape->setOffsetX(100);
    $shape->setOffsetY(200);
    $shape->getActiveParagraph()->setLineSpacing(200);

    $shape->createTextRun('RichText with');
    $shape->createBreak();
    $shape->createTextRun('Line Spacing 200');

    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text) with line spacing (300)' . EOL;
    $shape = $currentSlide->createRichTextShape();
    $shape->setHeight(100);
    $shape->setWidth(400);
    $shape->setOffsetX(100);
    $shape->setOffsetY(300);
    $shape->getActiveParagraph()->setLineSpacing(300);

    $shape->createTextRun('RichText with');
    $shape->createBreak();
    $shape->createTextRun('Line Spacing 300');
}

function fnSlideRichTextShadow(PhpPresentation $objPHPPresentation)
{
    // Create templated slide
    echo date('H:i:s') . ' Create templated slide' . EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text) with shadow' . EOL;
    $shape = $currentSlide->createRichTextShape();
    $shape->setHeight(100);
    $shape->setWidth(400);
    $shape->setOffsetX(100);
    $shape->setOffsetY(100);
    $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    $shape->getShadow()->setVisible(true)->setAlpha(75)->setBlurRadius(2)->setDirection(45);

    $textRun = $shape->createTextRun('RichText with shadow');
    $textRun->getFont()->setColor(new Color('FF000000'));
}

function fnSlideRichTextList(PhpPresentation $objPHPPresentation)
{
    // Create templated slide
    echo date('H:i:s') . ' Create templated slide' . EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text) with list with red bullet' . EOL;
    $shape = $currentSlide->createRichTextShape();
    $shape->setHeight(100);
    $shape->setWidth(400);
    $shape->setOffsetX(100);
    $shape->setOffsetY(100);
    $shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET)->setBulletColor(new Color(Color::COLOR_RED));

    $shape->createTextRun('Alpha');
    $shape->createParagraph()->createTextRun('Beta');
    $shape->createParagraph()->createTextRun('Delta');
    $shape->createParagraph()->createTextRun('Epsilon');
}

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$oProperties = $objPHPPresentation->getDocumentProperties();
$oProperties->setCreator('PHPOffice')
    ->setLastModifiedBy('PHPPresentation Team')
    ->setTitle('Sample 11 Title')
    ->setSubject('Sample 11 Subject')
    ->setDescription('Sample 11 Description')
    ->setKeywords('office 2007 openxml libreoffice odt php')
    ->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide' . EOL;
$objPHPPresentation->removeSlideByIndex(0);

fnSlideRichText($objPHPPresentation);
fnSlideRichTextLineSpacing($objPHPPresentation);
fnSlideRichTextShadow($objPHPPresentation);
fnSlideRichTextList($objPHPPresentation);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
