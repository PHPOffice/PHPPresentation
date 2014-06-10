<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Bullet;
use PhpOffice\PhpPowerpoint\Style\Color;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new \PhpOffice\PhpPowerpoint\PHPPowerPoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('Maarten Balliauw')
                                  ->setLastModifiedBy('Maarten Balliauw')
                                  ->setTitle('Office 2007 PPTX Test Document')
                                  ->setSubject('Office 2007 PPTX Test Document')
                                  ->setDescription('Test document for Office 2007 PPTX, generated using PHP classes.')
                                  ->setKeywords('office 2007 openxml php')
                                  ->setCategory('Test result file');

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPowerPoint->removeSlideByIndex(0);

// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function


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
$textRun->getFont()->setColor( new Color( 'FFFFFFFF' ) );

$shape->createBreak();

$textRun = $shape->createTextRun('PHPPowerPoint');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(60);
$textRun->getFont()->setColor( new Color( 'FFFFFFFF' ) );


// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('What is PHPPowerPoint?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor( new Color( 'FFFFFFFF' ) );

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape()
      ->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(100);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                            ->setMarginLeft(25)
                                            ->setIndent(-25);
$shape->getActiveParagraph()->getFont()->setSize(36)
                                       ->setColor(new Color('FFFFFFFF'));
$shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);

$shape->createTextRun('A class library');
$shape->createParagraph()->createTextRun('Written in PHP');
$shape->createParagraph()->createTextRun('Representing a presentation');
$shape->createParagraph()->createTextRun('Supports writing to different file formats');


// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape()
      ->setHeight(100)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('What\'s the point?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor(new Color('FFFFFFFF'));

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(100);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
                                            ->setMarginLeft(25)
                                            ->setIndent(-25);
$shape->getActiveParagraph()->getFont()->setSize(36)
                                       ->setColor(new Color('FFFFFFFF'));
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
$shape->createTextRun('PowerPoint 2007');
$shape->createParagraph()->createTextRun('Serialized');
$shape->createParagraph()->createTextRun('... (more to come) ...');


// Create templated slide
echo date('H:i:s') . ' Create templated slide'.EOL;
$currentSlide = createTemplatedSlide($objPHPPowerPoint); // local function

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Need more info?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor( new Color( 'FFFFFFFF' ) );

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(100);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Check the project site on CodePlex:');
$textRun->getFont()->setSize(36)
                   ->setColor( new Color( 'FFFFFFFF' ) );

$shape->createBreak();

$textRun = $shape->createTextRun('http://phppowerpoint.codeplex.com');
$textRun->getFont()->setSize(36)
                   ->setColor( new Color( 'FFFFFFFF' ) );
$textRun->getHyperlink()->setUrl('http://phppowerpoint.codeplex.com')
                        ->setTooltip('PHPPowerPoint');

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}

/**
 * Creates a templated slide
 *
 * @param PHPPowerPoint $objPHPPowerPoint
 * @return PHPPowerPoint_Slide
 */
function createTemplatedSlide(PhpPowerpoint $objPHPPowerPoint){
    // Create slide
    $slide = $objPHPPowerPoint->createSlide();

    // Add background image
    $slide->createDrawingShape()
          ->setName('Background')
          ->setDescription('Background')
          ->setPath('./resources/realdolmen_bg.jpg')
          ->setWidth(950)
          ->setHeight(720)
          ->setOffsetX(0)
          ->setOffsetY(0);

    // Add logo
    $slide->createDrawingShape()
          ->setName('PHPPowerPoint logo')
          ->setDescription('PHPPowerPoint logo')
          ->setPath('./resources/phppowerpoint_logo.gif')
          ->setHeight(40)
          ->setOffsetX(10)
          ->setOffsetY(720 - 10 - 40);

    // Return slide
    return $slide;
}
