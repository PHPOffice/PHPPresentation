<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Audio;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Animation;


// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
echo date('H:i:s') . ' Create slide'.EOL;

// new Presentation
$objPHPPresentation = new PhpPresentation();

// Create slide
$oSlide1 = $objPHPPresentation->getActiveSlide();


addAudioToSlide($oSlide1, './resources/Rick Astley - Never Gonna Give You Up.mp3', 20,10);

addAudioToSlide($oSlide1, './resources/a-ha - Take On Me (Official Music Video).mp3', 20,100);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}



function addAudioToSlide($slide, $path, $x,$y) {
	// Create a Audio
	$oAudio = new Audio();
	// settings
	$oAudio->setPath($path)
	->setOffsetX($x)
	->setOffsetY($y);

	// add to slide
	$slide->addShape($oAudio);
	
	// Create Animation
	$oAnimation = new Animation();
	// Add Audio to animation
	$oAnimation->addShape($oAudio);
	//Add animation to Slide
	$slide->addAnimation($oAnimation);
}