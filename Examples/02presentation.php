<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);

/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . '../Classes/');

/** PHPPowerPoint */
include 'PHPPowerPoint.php';

if(php_sapi_name() == 'cli' && empty($_SERVER['REMOTE_ADDR'])) {
    define('EOL', PHP_EOL);
}
else {
    define('EOL', '<br />');
}

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PHPPowerPoint();

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
$shape->getActiveParagraph()->getAlignment()->setHorizontal( PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Introduction to');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(28);
$textRun->getFont()->setColor( new PHPPowerPoint_Style_Color( 'FFFFFFFF' ) );

$shape->createBreak();

$textRun = $shape->createTextRun('PHPPowerPoint');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(60);
$textRun->getFont()->setColor( new PHPPowerPoint_Style_Color( 'FFFFFFFF' ) );


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
$shape->getActiveParagraph()->getAlignment()->setHorizontal( PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('What is PHPPowerPoint?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor( new PHPPowerPoint_Style_Color( 'FFFFFFFF' ) );

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape()
      ->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(100);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT)
                                            ->setMarginLeft(25)
                                            ->setIndent(-25);
$shape->getActiveParagraph()->getFont()->setSize(36)
                                       ->setColor(new PHPPowerPoint_Style_Color('FFFFFFFF'));
$shape->getActiveParagraph()->getBulletStyle()->setBulletType(PHPPowerPoint_Style_Bullet::TYPE_BULLET);

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
$shape->getActiveParagraph()->getAlignment()->setHorizontal(PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('What\'s the point?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor(new PHPPowerPoint_Style_Color('FFFFFFFF'));

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(100);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT)
                                            ->setMarginLeft(25)
                                            ->setIndent(-25);
$shape->getActiveParagraph()->getFont()->setSize(36)
                                       ->setColor(new PHPPowerPoint_Style_Color('FFFFFFFF'));
$shape->getActiveParagraph()->getBulletStyle()->setBulletType(PHPPowerPoint_Style_Bullet::TYPE_BULLET);

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
$shape->getActiveParagraph()->getAlignment()->setHorizontal( PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Need more info?');
$textRun->getFont()->setBold(true)
                   ->setSize(48)
                   ->setColor( new PHPPowerPoint_Style_Color( 'FFFFFFFF' ) );

// Create a shape (text)
echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setHeight(600)
      ->setWidth(930)
      ->setOffsetX(10)
      ->setOffsetY(100);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT );

$textRun = $shape->createTextRun('Check the project site on CodePlex:');
$textRun->getFont()->setSize(36)
                   ->setColor( new PHPPowerPoint_Style_Color( 'FFFFFFFF' ) );

$shape->createBreak();

$textRun = $shape->createTextRun('http://phppowerpoint.codeplex.com');
$textRun->getFont()->setSize(36)
                   ->setColor( new PHPPowerPoint_Style_Color( 'FFFFFFFF' ) );
$textRun->getHyperlink()->setUrl('http://phppowerpoint.codeplex.com')
                        ->setTooltip('PHPPowerPoint');

// Save files
$basename = basename(__FILE__, '.php');
$formats = array('PowerPoint2007' => 'pptx', 'ODPresentation' => 'odp');
foreach ($formats as $format => $extension) {
    echo date('H:i:s') . " Write to {$format} format".EOL;
    $objWriter = PHPPowerPoint_IOFactory::createWriter($objPHPPowerPoint, $format);
    $objWriter->save("results/{$basename}.{$extension}");
}

// Echo memory peak usage
echo date('H:i:s') . ' Peak memory usage: ' . (memory_get_peak_usage(true) / 1024 / 1024) . ' MB'.EOL;

// Echo done
echo date('H:i:s') . ' Done writing file.'.EOL;

/**
 * Creates a templated slide
 *
 * @param PHPPowerPoint $objPHPPowerPoint
 * @return PHPPowerPoint_Slide
 */
function createTemplatedSlide(PHPPowerPoint $objPHPPowerPoint){
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
