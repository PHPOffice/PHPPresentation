<?php
include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
$objPHPPresentation = new PhpPresentation();
// Set properties
echo date('H:i:s') . ' Set properties' . EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
    ->setLastModifiedBy('PHPPresentation Team')
    ->setTitle('Sample 19 SlideMaster')
    ->setSubject('Sample 19 Subject')
    ->setDescription('Sample 19 Description')
    ->setKeywords('office 2007 openxml libreoffice odt php')
    ->setCategory('Sample Category');

// Create slide
echo date('H:i:s') . ' Create slide' . EOL;
$currentSlide = $objPHPPresentation->getActiveSlide();

// Create a master layout
echo date('H:i:s') . ' Create masterslide layout' . EOL;
// Some decorative lines
$oMasterSlide = $objPHPPresentation->getAllMasterSlides()[0];
$shape = $oMasterSlide->createLineShape(0, 670, 960, 670)->getBorder()->setColor(new Color(Color::COLOR_RED))->setLineWidth(2);
$shape = $oMasterSlide->createLineShape(0, 672, 960, 672)->getBorder()->setColor(new Color(Color::COLOR_WHITE))->setLineWidth(2);
$shape = $oMasterSlide->createLineShape(0, 674, 960, 674)->getBorder()->setColor(new Color(Color::COLOR_DARKBLUE))->setLineWidth(2);
// Title placeholder
$shape = $oMasterSlide->createRichTextShape();
$shape->setWidthAndHeight(960, 80)->setOffsetX(0)->setOffsetY(60);
$shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
$shape->getActiveParagraph()->getAlignment()
    ->setHorizontal(Alignment::HORIZONTAL_RIGHT)->setMarginLeft(200)->setMarginRight(50)
    ->setVertical(Alignment::VERTICAL_CENTER);
$shape->getShadow()->setVisible(true)
    ->setDirection(90)
    ->setDistance(10);
$shape->setAutoFit(RichText::AUTOFIT_NORMAL);
$textRun = $shape->createTextRun('Titel');
$textRun->getFont()->setBold(true)->setSize(30)->setColor(new Color(Color::COLOR_WHITE));
$shape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_TITLE));
// Date placeholder
$shape = $oMasterSlide->createRichTextShape();
$shape->setWidthAndHeight(140, 38)->setOffsetX(50)->setOffsetY(680);
$shape->getActiveParagraph()->getAlignment()
    ->setHorizontal(Alignment::HORIZONTAL_LEFT)
    ->setVertical(Alignment::VERTICAL_BASE);
$shape->setAutoFit(RichText::AUTOFIT_NORMAL);
$textRun = $shape->createTextRun('01-02-2000')->getFont()->setSize(18);
$shape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_DATETIME))->getPlaceholder()->setIdx(10);
// Footer placeholder
$shape = $oMasterSlide->createRichTextShape();
$shape->setWidthAndHeight(468, 38)->setOffsetX(246)->setOffsetY(680);
$shape->getActiveParagraph()->getAlignment()
    ->setHorizontal(Alignment::HORIZONTAL_CENTER)
    ->setVertical(Alignment::VERTICAL_BASE);
$shape->setAutoFit(RichText::AUTOFIT_NORMAL);
$textRun = $shape->createTextRun('Placeholder for Footer')->getFont()->setSize(18);
$shape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_FOOTER))->getPlaceholder()->setIdx(11);
// Slidenumber placeholder
$shape = $oMasterSlide->createRichTextShape();
$shape->setWidthAndHeight(140, 38)->setOffsetX(770)->setOffsetY(680);
$shape->getActiveParagraph()->getAlignment()
    ->setHorizontal(Alignment::HORIZONTAL_RIGHT)
    ->setVertical(Alignment::VERTICAL_BASE);
$shape->setAutoFit(RichText::AUTOFIT_NORMAL);
$textRun = $shape->createTextRun('')->getFont()->setSize(18);
$shape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_SLIDENUM))->getPlaceholder()->setIdx(12);

// Create a shape (drawing)
echo date('H:i:s') . ' Create a shape (drawing)' . EOL;
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPresentation logo')
    ->setDescription('PHPPresentation logo')
    ->setPath('./resources/phppowerpoint_logo.gif')
    ->setHeight(36)
    ->setOffsetX(10)
    ->setOffsetY(10);
$shape->getShadow()->setVisible(true)
    ->setDirection(45)
    ->setDistance(10);
$shape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

// Create a shape (text) linked to a PlaceHolder
echo date('H:i:s') . ' Create a shape (rich text)' . EOL;
$shape = $currentSlide->createRichTextShape();
$shape->setWidthAndHeight(960, 80)->setOffsetX(0)->setOffsetY(60);
$shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
$shape->getActiveParagraph()->getAlignment()
    ->setHorizontal(Alignment::HORIZONTAL_RIGHT)->setMarginLeft(200)->setMarginRight(50)
    ->setVertical(Alignment::VERTICAL_CENTER);
$shape->setAutoFit(RichText::AUTOFIT_NORMAL);
$shape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_TITLE));

echo date('H:i:s') . ' Apply the shape collection of the slide master to the slide layout' . EOL;
$oSlideLayout = $oMasterSlide->getAllSlideLayouts()[0];
$oSlideLayout->setShapeCollection($oMasterSlide->getShapeCollection());
$currentSlide->setSlideLayout($oSlideLayout);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}