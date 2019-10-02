<?php

set_time_limit(10);

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Shape\RichText;

echo '<h2>ODPresentation</h2>';
$pptReader = IOFactory::createReader('ODPresentation');
$pptReader->setPassword('motdepasse');
$oPHPPresentation = $pptReader->load('resources/SamplePassword.odp');

$oTree = new PhpPptTree($oPHPPresentation);
echo $oTree->display();

echo '<h2>PowerPoint2007</h2>';
$pptReader = IOFactory::createReader('PowerPoint2007');
$pptReader->setPassword('motdepasse');
$oPHPPresentation = $pptReader->load('resources/SamplePassword.pptx');

$oTree = new PhpPptTree($oPHPPresentation);
echo $oTree->display();
if (!CLI) {
    include_once 'Sample_Footer.php';
}
