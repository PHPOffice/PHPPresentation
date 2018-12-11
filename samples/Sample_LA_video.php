<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
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


addVideoToSlide($oSlide1, './resources/sintel_trailer-480p.wmv', 20, 10, 850, 480);



// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
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