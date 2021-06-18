<?php

set_time_limit(10);

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\IOFactory;

$pptReader = IOFactory::createReader('PowerPoint97');
$oPHPPresentation = $pptReader->load('resources/Sample_12.ppt');

$oTree = new PhpPptTree($oPHPPresentation);
echo $oTree->display();
if (!CLI) {
    include_once 'Sample_Footer.php';
}
