<?php

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;

include_once 'Sample_Header.php';

$colorBlack = new Color( 'FF000000' );

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object'.EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPresentation Team')
                                  ->setTitle('Sample 02 Title')
                                  ->setSubject('Sample 02 Subject')
                                  ->setDescription('Sample 02 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPresentation->removeSlideByIndex(0);

// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation); // local function

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(200);
$shape->setWidth(600);
$shape->setOffsetX(10);
$shape->setOffsetY(400);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Introduction to');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(28);
$textRun->getFont()->setColor($colorBlack);

$shape->createBreak();

$textRun = $shape->createTextRun('PHPPresentation');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(60);
$textRun->getFont()->setColor($colorBlack);


// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation); // local function

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(50);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

$textRun = $shape->createTextRun('What is PHPPresentation?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor($colorBlack);

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape()
      ->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(130);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                            ->setMarginLeft(25)
                                            ->setIndent(-25);
$shape->getActiveParagraph()->getFont()->setSize(36)
                                       ->setColor($colorBlack);
$shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);

$shape->createTextRun('A class library');
$shape->createParagraph()->createTextRun('Written in PHP');
$shape->createParagraph()->createTextRun('Representing a presentation');
$shape->createParagraph()->createTextRun('Supports writing to different file formats');


// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation); // local function

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape()
      ->setHeight(100)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(50);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('What\'s the point?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor($colorBlack);

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(130);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                            ->setMarginLeft(25)
                                            ->setIndent(-25);
$shape->getActiveParagraph()->getFont()->setSize(36)
                                       ->setColor($colorBlack);
$shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);

$shape->createTextRun('Generate slide decks');
$shape->createParagraph()->getAlignment()->setLevel(1)
                                         ->setMarginLeft(75)
                                         ->setIndent(-25);
$shape->createTextRun('Represent business data');
$shape->createParagraph()->createTextRun('Show a family slide show');
$shape->createParagraph()->createTextRun('...');

$shape->createParagraph()->getAlignment()->setLevel(0)
                                         ->setMarginLeft(25)
                                         ->setIndent(-25);
$shape->createTextRun('Export these to different formats');
$shape->createParagraph()->getAlignment()->setLevel(1)
                                         ->setMarginLeft(75)
                                         ->setIndent(-25);
$shape->createTextRun('PHPPresentation 2007');
$shape->createParagraph()->createTextRun('Serialized');
$shape->createParagraph()->createTextRun('... (more to come) ...');


// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPresentation); // local function

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(50);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Need more info?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor($colorBlack);

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(130);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Check the project site on GitHub:');
$textRun->getFont()->setSize(36)
                   ->setColor($colorBlack);

$shape->createBreak();

$textRun = $shape->createTextRun('https://github.com/PHPOffice/PHPPresentation/');
$textRun->getFont()->setSize(32)
                   ->setColor($colorBlack);
$textRun->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')
                        ->setTooltip('PHPPresentation');

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
