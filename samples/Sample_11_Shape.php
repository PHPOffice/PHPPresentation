<?php

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Bullet;
use PhpOffice\PhpPowerpoint\Style\Color;

include_once 'Sample_Header.php';

function fnSlideRichText(PhpPowerpoint $objPHPPowerPoint) {
	// Create templated slide
	echo date('H:i:s') . ' Create templated slide'.EOL;
	$currentSlide = createTemplatedSlide($objPHPPowerPoint);
	
	// Create a shape (text)
	echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
	$shape = $currentSlide->createRichTextShape();
	$shape->setHeight(100);
	$shape->setWidth(600);
	$shape->setOffsetX(100);
	$shape->setOffsetY(100);
	$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
	
	$textRun = $shape->createTextRun('RichText with');
	$textRun->getFont()->setBold(true);
	$textRun->getFont()->setSize(28);
	$textRun->getFont()->setColor(new Color( 'FF000000' ));
	
	$shape->createBreak();
	
	$textRun = $shape->createTextRun('Multiline');
	$textRun->getFont()->setBold(true);
	$textRun->getFont()->setSize(60);
	$textRun->getFont()->setColor(new Color( 'FF000000' ));
}

function fnSlideRichTextShadow(PhpPowerpoint $objPHPPowerPoint) {
    // Create templated slide
    echo date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPowerPoint);

    // Create a shape (text)
    echo date('H:i:s') . ' Create a shape (rich text) with shadow'.EOL;
    $shape = $currentSlide->createRichTextShape();
	$shape->setHeight(100);
	$shape->setWidth(400);
	$shape->setOffsetX(100);
	$shape->setOffsetY(100);
    $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
    $shape->getShadow()->setVisible(true)->setAlpha(75)->setBlurRadius(2)->setDirection(45);

    $textRun = $shape->createTextRun('RichText with shadow');
    $textRun->getFont()->setColor(new Color( 'FF000000' ));
}

// Create new PHPPowerPoint object
echo date('H:i:s') . ' Create new PHPPowerPoint object'.EOL;
$objPHPPowerPoint = new PhpPowerpoint();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$oProperties = $objPHPPowerPoint->getProperties();
$oProperties->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPowerPoint Team')
            ->setTitle('Sample 11 Title')
            ->setSubject('Sample 11 Subject')
            ->setDescription('Sample 11 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPowerPoint->removeSlideByIndex(0);

fnSlideRichText($objPHPPowerPoint);
fnSlideRichTextShadow($objPHPPowerPoint);

// Save file
echo write($objPHPPowerPoint, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}
