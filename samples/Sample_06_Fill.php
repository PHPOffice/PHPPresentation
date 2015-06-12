<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Fill;

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object' . EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPowerPoint->getProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPowerPoint Team')
                                  ->setTitle('Sample 01 Title')
                                  ->setSubject('Sample 01 Subject')
                                  ->setDescription('Sample 01 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide'.EOL;
$currentSlide = $objPHPPowerPoint->getActiveSlide();


for($inc = 1 ; $inc <= 4 ; $inc++){
    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
    $shape = $currentSlide->createRichTextShape()
                          ->setHeight(200)
                          ->setWidth(300);
    if($inc == 1 || $inc == 3){
        $shape->setOffsetX(10);
    } else {
        $shape->setOffsetX(320);
    }
    if($inc == 1 || $inc == 2){
        $shape->setOffsetY(10);
    } else {
        $shape->setOffsetY(220);
    }
    $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
    
    switch ($inc) {
        case 1 :
            $shape->getFill()->setFillType(Fill::FILL_NONE);
            break;
        case 2 :
            $shape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setRotation(90)->setStartColor(new Color( 'FF4672A8' ))->setEndColor(new Color( 'FF000000' ));
            break;
        case 3 :
            $shape->getFill()->setFillType(Fill::FILL_GRADIENT_PATH)->setRotation(90)->setStartColor(new Color( 'FF4672A8' ))->setEndColor(new Color( 'FF000000' ));
            break;
        case 4 :
            $shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color( 'FF4672A8' ))->setEndColor(new Color( 'FF4672A8' ));
            break;
    }
    
    $textRun = $shape->createTextRun('Use PHPPowerPoint!');
    $textRun->getFont()->setBold(true)
                       ->setSize(30)
                       ->setColor( new Color('FFE06B20') );
}

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
