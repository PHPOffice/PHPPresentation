<?php

set_time_limit(10);

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Group;

$pptReader = PhpOffice\PhpPresentation\IOFactory::createReader('PowerPoint97');
$oPHPPresentation = $pptReader->load('resources/Sample_12.ppt');

$oTree = new PhpPptTree($oPHPPresentation);
echo $oTree->display();
if (!CLI) {
	include_once 'Sample_Footer.php';
}
