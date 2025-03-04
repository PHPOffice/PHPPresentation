<?php

set_time_limit(10);

include_once __DIR__ . '/Sample_PhpPptTree.php';
include_once __DIR__ . '/Sample_Header.php';

use PhpOffice\PhpPresentation\IOFactory;

$pptReader = IOFactory::createReader('PowerPoint2007');
$oPHPPresentation = $pptReader->load(__DIR__ . '/resources/Sample_12.pptx');

if (!CLI) {
    $oTree = new Sample_PhpPptTree($oPHPPresentation);
    echo $oTree->display();

    include_once 'Sample_Footer.php';
}
