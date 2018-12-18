<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Audio;
use PhpOffice\PhpPresentation\Shape\Video;
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
addVideoToSlide($oSlide1, './resources/sintel_trailer-480p.wmv', 20, 100, 850, 480);


$oSlide2 = $objPHPPresentation->createSlide();
addAudioToSlide($oSlide2, './resources/Rick Astley - Never Gonna Give You Up.mp3', 20,500);
addVideoToSlide($oSlide2, './resources/sintel_trailer-480p.wmv', 20, 10, 850, 480);



$oSlide3 = $objPHPPresentation->createSlide();
addAudioToSlide($oSlide3, './resources/Rick Astley - Never Gonna Give You Up.mp3', 20,10);
addAudioToSlide($oSlide3, './resources/Rick Astley - Never Gonna Give You Up.mp3', 20,100);
addAudioToSlide($oSlide3, './resources/Rick Astley - Never Gonna Give You Up.mp3', 20,200);
addAudioToSlide($oSlide3, './resources/Rick Astley - Never Gonna Give You Up.mp3', 20,300);

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

function addVideoToSlide($slide, $path, $x, $y, $width, $height) {
	// Create a Video
	$oVideo = new Video();
	// settings
	$oVideo->setPath($path)
	->setHeight($height)
	->setWidth($width)
	->setOffsetX($x)
	->setOffsetY($y);
	// add to slide
	$slide->addShape($oVideo);
	
	// Create Animation
	$oAnimation = new Animation();
	// Add video to animation
	$oAnimation->addShape($oVideo);
	//Add animation to Slide
	$slide->addAnimation($oAnimation);
}