<?php

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\ODPresentation;
use PhpOffice\Common\Drawing;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 */
class ContentTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    public function testDrawingWithHyperlink()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');
    
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/office:event-listeners/presentation:event-listener';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'https://github.com/PHPOffice/PHPPresentation/');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testDrawingShapeFill()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        $element = '/office:document-content/office:automatic-styles/style:style/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'none');

        $oColor = new Color(Color::COLOR_DARKRED);
        $oColor->setAlpha(rand(0, 100));
        $oShape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($oColor);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:automatic-styles/style:style/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:fill-color', '#');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:fill-color', $oColor->getRGB());
    }

    public function testComment()
    {
        $expectedName = 'Name';
        $expectedText = 'Text';

        $oAuthor = new Comment\Author();
        $oAuthor->setName($expectedName);
        $oComment = new Comment();
        $oComment->setAuthor($oAuthor);
        $oComment->setText($expectedText);
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $element = '/office:document-content';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'xmlns:officeooo');
        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation/dc:creator';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementEquals('content.xml', $element, $expectedName);
        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation/text:p';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementEquals('content.xml', $element, $expectedText);
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testCommentWithoutAuthor()
    {
        $oComment = new Comment();
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'dc:creator');
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testFillGradientLinearRichText()
    {
        $oShape = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oShape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FFFF7700'))->setEndColor(new Color('FFFFFFFF'));

        $element = '/office:document-styles/office:styles/draw:gradient';
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:name', 'gradient_' . $oShape->getFill()->getHashCode());

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'gradient');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-gradient-name', 'gradient_' . $oShape->getFill()->getHashCode());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testFillSolidRichText()
    {
        $oShape = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oShape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF4672A8'));

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#' . $oShape->getFill()->getStartColor()->getRGB());
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#' . $oShape->getFill()->getEndColor()->getRGB());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testGroup()
    {
        $oShapeGroup = $this->oPresentation->getActiveSlide()->createGroup();
        $oShape = $oShapeGroup->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');
    
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:g';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:g/draw:frame/office:event-listeners/presentation:event-listener';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
    
    public function testList()
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');
        
        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testInnerList()
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)->setMarginLeft(25)->setIndent(-25);
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->getAlignment()->setLevel(1)->setMarginLeft(75)->setIndent(-25);
        $oRichText->createTextRun('Alpha.Alpha');
        $oRichText->createParagraph()->createTextRun('Alpha.Beta');
        $oRichText->createParagraph()->createTextRun('Alpha.Delta');
        
        $oRichText->createParagraph()->getAlignment()->setLevel(0)->setMarginLeft(25)->setIndent(-25);
        $oRichText->createTextRun('Beta');
        $oRichText->createParagraph()->getAlignment()->setLevel(1)->setMarginLeft(75)->setIndent(-25);
        $oRichText->createTextRun('Beta.Alpha');
        $oRichText->createParagraph()->createTextRun('Beta.Beta');
        $oRichText->createParagraph()->createTextRun('Beta.Delta');
        
        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:list/text:list-item/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testParagraphRichText()
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('Alpha');
        $oRichText->createBreak();
        $oRichText->createText('Beta');
        $oRichText->createBreak();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:line-break';
        $this->assertZipXmlElementExists('content.xml', $element);
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:a';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'http://www.google.fr');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testListWithRichText()
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRun = $oRichText->createTextRun('Alpha');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
        $oRichText->createBreak();
        $oRichText->createTextRun('Beta');
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:a';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:line-break';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testMedia()
    {
        $expectedName = 'MyName';
        $expectedWidth = mt_rand(1, 100);
        $expectedHeight = mt_rand(1, 100);
        $expectedX = mt_rand(1, 100);
        $expectedY = mt_rand(1, 100);

        $oMedia = new Media();
        $oMedia->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/videos/sintel_trailer-480p.ogv')
            ->setName($expectedName)
            ->setResizeProportional(false)
            ->setHeight($expectedHeight)
            ->setWidth($expectedWidth)
            ->setOffsetX($expectedX)
            ->setOffsetY($expectedY);
        $this->oPresentation->getActiveSlide()->addShape($oMedia);

        $expectedWidth = Text::numberFormat(CommonDrawing::pixelsToCentimeters($expectedWidth), 3) . 'cm';
        $expectedHeight = Text::numberFormat(CommonDrawing::pixelsToCentimeters($expectedHeight), 3) . 'cm';
        $expectedX = Text::numberFormat(CommonDrawing::pixelsToCentimeters($expectedX), 3) . 'cm';
        $expectedY = Text::numberFormat(CommonDrawing::pixelsToCentimeters($expectedY), 3) . 'cm';

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:name', $expectedName);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:width', $expectedWidth);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:height', $expectedHeight);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:x', $expectedX);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:y', $expectedY);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:plugin';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:mime-type', 'application/vnd.sun.star.media');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:mime-type', 'application/vnd.sun.star.media');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'xlink:href', 'Pictures/');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'xlink:href', 'ogv');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testNote()
    {
        $oNote = $this->oPresentation->getActiveSlide()->getNote();
        $oRichText = $oNote->createRichTextShape()->setHeight(300)->setWidth(600);
        $oRichText->createTextRun('testNote');

        $element = '/office:document-content/office:body/office:presentation/draw:page/presentation:notes';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/presentation:notes/draw:frame/draw:text-box/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextAutoShrink()
    {
        $oRichText1 = $this->oPresentation->getActiveSlide()->createRichTextShape();

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:auto-grow-height');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:auto-grow-width');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRichText1->setAutoShrinkHorizontal(false);
        $oRichText1->setAutoShrinkVertical(true);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-height');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-width');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-height', 'true');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-width', 'false');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRichText1->setAutoShrinkHorizontal(true);
        $oRichText1->setAutoShrinkVertical(false);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-height');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-width');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-height', 'false');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-width', 'true');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextRunLanguage()
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('MyText');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'T_' . $oRun->getHashCode() . '\']/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:language', 'en');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->setLanguage('de');
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:language', 'de');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextBorder()
    {
        $oRichText1 = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_NONE);
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'svg:stroke-width');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:stroke');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:stroke', 'none');
        $this->assertIsSchemaOpenDocumentValid('1.2');
        
        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_SINGLE);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:stroke-color', '#' . $oRichText1->getBorder()->getColor()->getRGB());
        $this->assertZipXmlAttributeExists('content.xml', $element, 'svg:stroke-width');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'svg:stroke-width', (string)number_format(CommonDrawing::pointsToCentimeters($oRichText1->getBorder()->getLineWidth()), 3, '.', ''));
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'svg:stroke-width', 'cm');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:stroke');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:stroke', 'solid');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:stroke-dash');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_DASH);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:stroke', 'dash');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:stroke-dash');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:stroke-dash', 'strokeDash_');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:stroke-dash', $oRichText1->getBorder()->getDashStyle());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
    
    public function testRichTextShadow()
    {
        $randAlpha = mt_rand(0, 100);
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->getShadow()->setVisible(true)->setAlpha($randAlpha)->setBlurRadius(2);
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        for ($inc = 0; $inc <= 360; $inc += 45) {
            $randDistance = mt_rand(0, 100);
            $oRichText->getShadow()->setDirection($inc)->setDistance($randDistance);

            $this->assertZipXmlElementExists('content.xml', $element);
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow', 'visible');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:mirror', 'none');
            // Opacity
            $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:shadow-opacity', (string)(100 - $randAlpha));
            $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:shadow-opacity', '%');
            // Color
            $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:shadow-color', '#');
            // X
            if ($inc == 90 || $inc == 270) {
                $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-x', '0cm');
            } else {
                if ($inc > 90 && $inc < 270) {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-x', '-' . Drawing::pixelsToCentimeters($randDistance) . 'cm');
                } else {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-x', Drawing::pixelsToCentimeters($randDistance) . 'cm');
                }
            }
            // Y
            if ($inc == 0 || $inc == 180 || $inc == 360) {
                $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-y', '0cm');
            } else {
                if (($inc > 0 && $inc < 180) || $inc == 360) {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-y', Drawing::pixelsToCentimeters($randDistance) . 'cm');
                } else {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-y', '-' . Drawing::pixelsToCentimeters($randDistance) . 'cm');
                }
            }
            $this->assertIsSchemaOpenDocumentValid('1.2');
            $this->resetPresentationFile();
        }
    }

    public function testSlideName()
    {
        $element = '/office:document-content/office:body/office:presentation/draw:page';

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:name');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->oPresentation->getActiveSlide()->setName('AAAA');
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:name');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:name', 'AAAA');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->oPresentation->getActiveSlide()->setName();
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:name');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleAlignment()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        $oRichText1->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $oRichText1->createTextRun('Run1');
        $oRichText2 = $oSlide->createRichTextShape();
        $oRichText2->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_DISTRIBUTED);
        $oRichText2->createTextRun('Run2');
        $oRichText3 = $oSlide->createRichTextShape();
        $oRichText3->getActiveParagraph()->getAlignment()->setHorizontal('AAAAA');
        $oRichText3->createTextRun('Run3');
        $oRichText4 = $oSlide->createRichTextShape();
        $oRichText4->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);
        $oRichText4->createTextRun('Run4');
        $oRichText5 = $oSlide->createRichTextShape();
        $oRichText5->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $oRichText5->createTextRun('Run5');
        $oRichText6 = $oSlide->createRichTextShape();
        $oRichText6->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $oRichText6->createTextRun('Run6');

        $p1HashCode = $oRichText1->getActiveParagraph()->getHashCode();
        $p2HashCode = $oRichText2->getActiveParagraph()->getHashCode();
        $p3HashCode = $oRichText3->getActiveParagraph()->getHashCode();
        $p4HashCode = $oRichText4->getActiveParagraph()->getHashCode();
        $p5HashCode = $oRichText5->getActiveParagraph()->getHashCode();
        $p6HashCode = $oRichText6->getActiveParagraph()->getHashCode();
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p1HashCode.'\']/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'center');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p2HashCode.'\']/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'justify');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p3HashCode.'\']/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'left');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p4HashCode.'\']/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'justify');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p5HashCode.'\']/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'left');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p6HashCode.'\']/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'right');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
    
    public function testStyleFont()
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setBold(true);
        
        $expectedHashCode = $oRun->getHashCode();
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'T_'.$expectedHashCode.'\']/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:font-weight', 'bold');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
    
    public function testTable()
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oShape->createRow();

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';
        $this->assertZipXmlElementExists('content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableEmpty()
    {
        $this->oPresentation->getActiveSlide()->createTableShape();

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';
        $this->assertZipXmlElementNotExists('content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
    
    public function testTableCellFill()
    {
        $oColor = new Color();
        $oColor->setRGB(Color::COLOR_BLUE);
        
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor($oColor);

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->setFill($oFill);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1r0c0\']';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:family', 'table-cell');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1r0c0\']/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:fill-color', '#');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:fill-color', $oColor->getRGB());

        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }
    
    public function testTableWithColspan()
    {
        $value = mt_rand(2, 100);

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape($value);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->setColSpan($value);

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'table:number-columns-spanned', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
    
    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/70
     */
    public function testTableWithHyperlink()
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oTextRun = $oCell->createTextRun('AAA');
        $oHyperlink = $oTextRun->getHyperlink();
        $oHyperlink->setUrl('https://github.com/PHPOffice/PHPPresentation/');
    
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span/text:a';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'https://github.com/PHPOffice/PHPPresentation/');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
    
    public function testTableWithText()
    {
        $oRun = new Run();
        $oRun->setText('Test');

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->addText($oRun);
        $oCell->createBreak();

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementEquals('content.xml', $element, 'Test');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span/text:line-break';
        $this->assertZipXmlElementExists('content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTransition()
    {
        $value = mt_rand(1000, 5000);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePage0\']/style:drawing-page-properties';

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:duration');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition = new Transition();
        $oTransition->setTimeTrigger(true, $value);
        $this->oPresentation->getActiveSlide()->setTransition($oTransition);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:duration');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'presentation:duration', 'PT');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'presentation:duration', 'S');
        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:duration', number_format($value / 1000, 6, '.', ''));
        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-type', 'automatic');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition->setSpeed(Transition::SPEED_FAST);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-speed', 'fast');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition->setSpeed(Transition::SPEED_MEDIUM);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-speed', 'medium');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition->setSpeed(Transition::SPEED_SLOW);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-speed', 'slow');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $rcTransition = new \ReflectionClass('PhpOffice\PhpPresentation\Slide\Transition');
        $arrayConstants = $rcTransition->getConstants();
        foreach ($arrayConstants as $key => $value) {
            if (strpos($key, 'TRANSITION_') !== 0) {
                continue;
            }
            $this->resetPresentationFile();
            $oTransition->setTransitionType($rcTransition->getConstant($key));
            $this->oPresentation->getActiveSlide()->setTransition($oTransition);
            switch ($key) {
                case 'TRANSITION_BLINDS_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'horizontal-stripes');
                    break;
                case 'TRANSITION_BLINDS_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'vertical-stripes');
                    break;
                case 'TRANSITION_CHECKER_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'horizontal-checkerboard');
                    break;
                case 'TRANSITION_CHECKER_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'vertical-checkerboard');
                    break;
                case 'TRANSITION_CIRCLE_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_CIRCLE_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_COMB_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_COMB_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_COVER_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-bottom');
                    break;
                case 'TRANSITION_COVER_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-left');
                    break;
                case 'TRANSITION_COVER_LEFT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-lowerleft');
                    break;
                case 'TRANSITION_COVER_LEFT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-upperleft');
                    break;
                case 'TRANSITION_COVER_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-right');
                    break;
                case 'TRANSITION_COVER_RIGHT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-lowerright');
                    break;
                case 'TRANSITION_COVER_RIGHT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-upperright');
                    break;
                case 'TRANSITION_COVER_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-top');
                    break;
                case 'TRANSITION_CUT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_DIAMOND':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_DISSOLVE':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'dissolve');
                    break;
                case 'TRANSITION_FADE':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-center');
                    break;
                case 'TRANSITION_NEWSFLASH':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_PLUS':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'close');
                    break;
                case 'TRANSITION_PULL_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-bottom');
                    break;
                case 'TRANSITION_PULL_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-left');
                    break;
                case 'TRANSITION_PULL_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-right');
                    break;
                case 'TRANSITION_PULL_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-top');
                    break;
                case 'TRANSITION_PUSH_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-bottom');
                    break;
                case 'TRANSITION_PUSH_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-left');
                    break;
                case 'TRANSITION_PUSH_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-right');
                    break;
                case 'TRANSITION_PUSH_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-top');
                    break;
                case 'TRANSITION_RANDOM':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'random');
                    break;
                case 'TRANSITION_RANDOMBAR_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'horizontal-lines');
                    break;
                case 'TRANSITION_RANDOMBAR_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'vertical-lines');
                    break;
                case 'TRANSITION_SPLIT_IN_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'close-horizontal');
                    break;
                case 'TRANSITION_SPLIT_OUT_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'open-horizontal');
                    break;
                case 'TRANSITION_SPLIT_IN_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'close-vertical');
                    break;
                case 'TRANSITION_SPLIT_OUT_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'open-vertical');
                    break;
                case 'TRANSITION_STRIPS_LEFT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_STRIPS_LEFT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_STRIPS_RIGHT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_STRIPS_RIGHT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_WEDGE':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_WIPE_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-bottom');
                    break;
                case 'TRANSITION_WIPE_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-left');
                    break;
                case 'TRANSITION_WIPE_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-right');
                    break;
                case 'TRANSITION_WIPE_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-top');
                    break;
                case 'TRANSITION_ZOOM_IN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
                case 'TRANSITION_ZOOM_OUT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');
                    break;
            }
            $this->assertIsSchemaOpenDocumentValid('1.2');
        }

        $oTransition->setTimeTrigger(false);
        $oTransition->setManualTrigger(true);
        $this->resetPresentationFile();
        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-type', 'manual');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testVisibility()
    {
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePage0\']/style:drawing-page-properties';

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:visibility');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->setIsVisible(false);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:visibility');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'presentation:visibility', 'hidden');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }
}
