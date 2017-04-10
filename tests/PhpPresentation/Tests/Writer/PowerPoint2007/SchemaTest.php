<?php

namespace PhpOffice\PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Animation;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Media;
use Symfony\Component\Finder\SplFileInfo;

class SchemaTest extends PhpPresentationTestCase
{
    private $internalErrors;

    public function setUp()
    {
        parent::setUp();

        libxml_clear_errors();
        $this->internalErrors = libxml_use_internal_errors(true);
    }

    public function tearDown()
    {
        parent::tearDown();

        libxml_use_internal_errors($this->internalErrors);
    }

    /**
     * Test whether the generated XML validates against the Office Open XML File Formats schema
     *
     * @see http://www.ecma-international.org/publications/standards/Ecma-376.htm
     * @dataProvider pptProvider
     */
    public function testSchema(PhpPresentation $presentation)
    {
        $this->writePresentationFile($presentation, 'PowerPoint2007');

        // validate all XML files
        $path = realpath($this->workDirectory . '/ppt');
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($path));

        foreach ($iterator as $file) {
            /** @var SplFileInfo $file */
            if ($file->getExtension() !== "xml") {
                continue;
            }

            $fileName = str_replace('\\', '/', substr($file->getRealPath(), strlen($path) + 1));
            $dom = $this->getXmlDom('ppt/' . $fileName);
            $xmlSource = $dom->saveXML();

            // In the ISO/ECMA standard the namespace has changed from
            // http://schemas.openxmlformats.org/ to http://purl.oclc.org/ooxml/
            // We need to use the http://purl.oclc.org/ooxml/ namespace to validate
            // the xml against the current schema
            $xmlSource = str_replace(array(
                "http://schemas.openxmlformats.org/drawingml/2006/main",
                "http://schemas.openxmlformats.org/drawingml/2006/chart",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "http://schemas.openxmlformats.org/presentationml/2006/main",
            ), array(
                "http://purl.oclc.org/ooxml/drawingml/main",
                "http://purl.oclc.org/ooxml/drawingml/chart",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
                "http://purl.oclc.org/ooxml/presentationml/main",
            ), $xmlSource);

            $dom->loadXML($xmlSource);
            $dom->schemaValidate(__DIR__ . '/../../../../resources/schema/ooxml/pml.xsd');

            $error = libxml_get_last_error();
            if ($error) {
                $this->fail(sprintf("Validation error: %s in file %s on line %s\n%s", $error->message, $file, $error->line, $dom->saveXML()));
            }
        }
    }

    public function pptProvider()
    {
        return array(
            array($this->generatePresentation01()),
            array($this->generatePresentation02()),
        );
    }

    /**
     * Generates a ppt which contains different elements per slide i.e. shape, table, etc.
     *
     * @return PhpPresentation
     */
    private function generatePresentation01()
    {
        $objPHPPresentation = new PhpPresentation();

        $objPHPPresentation->getDocumentProperties()
            ->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 02 Title')
            ->setSubject('Sample 02 Subject')
            ->setDescription('Sample 02 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        $currentSlide = $objPHPPresentation->getActiveSlide();

        // text shape
        $shape = $currentSlide->createRichTextShape()
            ->setHeight(300)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(180);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
        $textRun->getFont()->setBold(true)
            ->setSize(60)
            ->setColor(new Color('FFE06B20'));

        // image
        $currentSlide = $objPHPPresentation->createSlide();
        $gdImage = @imagecreatetruecolor(140, 20);
        $textColor = imagecolorallocate($gdImage, 255, 255, 255);
        imagestring($gdImage, 1, 5, 5, 'Created with PHPPresentation', $textColor);

        $shape = new Drawing\Gd();
        $shape->setName('Sample image')
            ->setDescription('Sample image')
            ->setImageResource($gdImage)
            ->setRenderingFunction(Drawing\Gd::RENDERING_JPEG)
            ->setMimeType(Drawing\Gd::MIMETYPE_DEFAULT)
            ->setHeight(36)
            ->setOffsetX(10)
            ->setOffsetY(10);
        $currentSlide->addShape($shape);

        // table
        $currentSlide = $objPHPPresentation->createSlide();
        $shape = $currentSlide->createTableShape(3);
        $shape->setHeight(200);
        $shape->setWidth(600);
        $shape->setOffsetX(150);
        $shape->setOffsetY(300);

        $row = $shape->createRow();
        $row->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)
            ->setRotation(90)
            ->setStartColor(new Color('FFE06B20'))
            ->setEndColor(new Color('FFFFFFFF'));
        $cell = $row->nextCell();
        $cell->setColSpan(3);
        $cell->createTextRun('Title row')->getFont()->setBold(true)->setSize(16);
        $cell->getBorders()->getBottom()->setLineWidth(4)
            ->setLineStyle(Border::LINE_SINGLE)
            ->setDashStyle(Border::DASH_DASH);

        $row = $shape->createRow();
        $row->setHeight(20);
        $row->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)
            ->setRotation(90)
            ->setStartColor(new Color('FFE06B20'))
            ->setEndColor(new Color('FFFFFFFF'));
        $row->nextCell()->createTextRun('R1C1')->getFont()->setBold(true);
        $row->nextCell()->createTextRun('R1C2')->getFont()->setBold(true);
        $row->nextCell()->createTextRun('R1C3')->getFont()->setBold(true);

        foreach ($row->getCells() as $cell) {
            $cell->getBorders()->getTop()->setLineWidth(4)
                ->setLineStyle(Border::LINE_SINGLE)
                ->setDashStyle(Border::DASH_DASH);
        }

        // chart
        $currentSlide = $objPHPPresentation->createSlide();

        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

        $seriesData = array(
            'Monday' => 12,
            'Tuesday' => 15,
            'Wednesday' => 13,
            'Thursday' => 17,
            'Friday' => 14,
            'Saturday' => 9,
            'Sunday' => 7
        );

        $lineChart = new Line();
        $series = new Series('Downloads', $seriesData);
        $series->setShowSeriesName(true);
        $series->setShowValue(true);
        $lineChart->addSeries($series);

        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getLegend()->getFont()->setItalic(true);

        // fill
        $currentSlide = $objPHPPresentation->createSlide();

        for ($inc = 1; $inc <= 4; $inc++) {
            // Create a shape (text)
            $shape = $currentSlide->createRichTextShape()
                ->setHeight(200)
                ->setWidth(300);
            if ($inc == 1 || $inc == 3) {
                $shape->setOffsetX(10);
            } else {
                $shape->setOffsetX(320);
            }
            if ($inc == 1 || $inc == 2) {
                $shape->setOffsetY(10);
            } else {
                $shape->setOffsetY(220);
            }
            $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

            switch ($inc) {
                case 1:
                    $shape->getFill()->setFillType(Fill::FILL_NONE);
                    break;
                case 2:
                    $shape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF000000'));
                    break;
                case 3:
                    $shape->getFill()->setFillType(Fill::FILL_GRADIENT_PATH)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF000000'));
                    break;
                case 4:
                    $shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF4672A8'));
                    break;
            }

            $textRun = $shape->createTextRun('Use PHPPresentation!');
            $textRun->getFont()->setBold(true)
                ->setSize(30)
                ->setColor(new Color('FFE06B20'));
        }

        // slide note
        $currentSlide = $objPHPPresentation->createSlide();

        $shape = $currentSlide->createDrawingShape();
        $shape->setName('PHPPresentation logo')
            ->setDescription('PHPPresentation logo')
            ->setPath(__DIR__ . '/../../../../../samples/resources/phppowerpoint_logo.gif')
            ->setHeight(36)
            ->setOffsetX(10)
            ->setOffsetY(10);
        $shape->getShadow()->setVisible(true)
            ->setDirection(45)
            ->setDistance(10);
        $shape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

        $shape = $currentSlide->createRichTextShape()
            ->setHeight(300)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(180);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
        $textRun->getFont()->setBold(true)
            ->setSize(60)
            ->setColor(new Color('FFE06B20'));

        $oNote = $currentSlide->getNote();
        $oLayout = $objPHPPresentation->getLayout();
        $oRichText = $oNote->createRichTextShape()
            ->setHeight($oLayout->getCY($oLayout::UNIT_PIXEL))
            ->setWidth($oLayout->getCX($oLayout::UNIT_PIXEL))
            ->setOffsetX(170)
            ->setOffsetY(180);
        $oRichText->createTextRun('A class library');
        $oRichText->createParagraph()->createTextRun('Written in PHP');
        $oRichText->createParagraph()->createTextRun('Representing a presentation');
        $oRichText->createParagraph()->createTextRun('Supports writing to different file formats');

        // transition
        $currentSlide = $objPHPPresentation->createSlide();

        $shapeDrawing = $currentSlide->createDrawingShape();
        $shapeDrawing->setName('PHPPresentation logo')
            ->setDescription('PHPPresentation logo')
            ->setPath(__DIR__ . '/../../../../../samples/resources/phppowerpoint_logo.gif')
            ->setHeight(36)
            ->setOffsetX(10)
            ->setOffsetY(10);
        $shapeDrawing->getShadow()->setVisible(true)
            ->setDirection(45)
            ->setDistance(10);
        $shapeDrawing->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

        $shapeRichText = $currentSlide->createRichTextShape()
            ->setHeight(300)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(180);
        $shapeRichText->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shapeRichText->createTextRun('Thank you for using PHPPresentation!');
        $textRun->getFont()->setBold(true)
            ->setSize(60)
            ->setColor(new Color('FFE06B20'));

        $oTransition = new Transition();
        $oTransition->setManualTrigger(false);
        $oTransition->setTimeTrigger(true, 4000);
        $oTransition->setTransitionType(Transition::TRANSITION_SPLIT_IN_VERTICAL);
        $currentSlide->setTransition($oTransition);

        // comment
        $currentSlide = $objPHPPresentation->createSlide();

        $oShapeDrawing = new Drawing\File();
        $oShapeDrawing->setName('PHPPresentation logo')
            ->setDescription('PHPPresentation logo')
            ->setPath(__DIR__ . '/../../../../../samples/resources/phppowerpoint_logo.gif')
            ->setHeight(36)
            ->setOffsetX(10)
            ->setOffsetY(10);
        $oShapeDrawing->getShadow()->setVisible(true)
            ->setDirection(45)
            ->setDistance(10);
        $oShapeDrawing->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

        $oShapeRichText = new RichText();
        $oShapeRichText->setHeight(300)
            ->setWidth(600)
            ->setOffsetX(170)
            ->setOffsetY(180);
        $oShapeRichText->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $oShapeRichText->createTextRun('Thank you for using PHPPresentation!');
        $textRun->getFont()->setBold(true)
            ->setSize(60)
            ->setColor(new Color('FFE06B20'));

        $currentSlide->addShape(clone $oShapeDrawing);
        $currentSlide->addShape(clone $oShapeRichText);

        $oAuthor = new \PhpOffice\PhpPresentation\Shape\Comment\Author();
        $oAuthor->setName('Progi1984');
        $oAuthor->setInitials('P');

        $oComment1 = new \PhpOffice\PhpPresentation\Shape\Comment();
        $oComment1->setText('Text A');
        $oComment1->setOffsetX(10);
        $oComment1->setOffsetY(55);
        $oComment1->setDate(time());
        $oComment1->setAuthor($oAuthor);
        $currentSlide->addShape($oComment1);

        $oComment2 = new \PhpOffice\PhpPresentation\Shape\Comment();
        $oComment2->setText('Text B');
        $oComment2->setOffsetX(170);
        $oComment2->setOffsetY(180);
        $oComment2->setDate(time());
        $oComment2->setAuthor($oAuthor);
        $currentSlide->addShape($oComment2);

        // animation
        $currentSlide = $objPHPPresentation->createSlide();

        $oDrawing1 = clone $oShapeDrawing;
        $oRichText1 = clone $oShapeRichText;

        $oSlide1 = $objPHPPresentation->getActiveSlide();
        $oSlide1->addShape($oDrawing1);
        $oSlide1->addShape($oRichText1);

        $oAnimation1 = new Animation();
        $oAnimation1->addShape($oDrawing1);
        $oSlide1->addAnimation($oAnimation1);

        $oAnimation2 = new Animation();
        $oAnimation2->addShape($oRichText1);
        $oSlide1->addAnimation($oAnimation2);

        $oDrawing2 = clone $oShapeDrawing;
        $oRichText2 = clone $oShapeRichText;

        $oSlide2 = $objPHPPresentation->createSlide();
        $oSlide2->addShape($oDrawing2);
        $oSlide2->addShape($oRichText2);

        $oAnimation4 = new Animation();
        $oAnimation4->addShape($oRichText2);
        $oSlide2->addAnimation($oAnimation4);

        $oAnimation3 = new Animation();
        $oAnimation3->addShape($oDrawing2);
        $oSlide2->addAnimation($oAnimation3);

        $oDrawing3 = clone $oShapeDrawing;
        $oRichText3 = clone $oShapeRichText;

        $currentSlide->addShape($oDrawing3);
        $currentSlide->addShape($oRichText3);

        $oAnimation5 = new Animation();
        $oAnimation5->addShape($oRichText3);
        $oAnimation5->addShape($oDrawing3);
        $currentSlide->addAnimation($oAnimation5);

        // @TODO add more complex elements

        return $objPHPPresentation;
    }

    /**
     * Generates a ppt containing placeholder in the master and the slide
     *
     * @return PhpPresentation
     */
    private function generatePresentation02()
    {
        $objPHPPresentation = new PhpPresentation();

        $objPHPPresentation->getDocumentProperties()
            ->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 02 Title')
            ->setSubject('Sample 02 Subject')
            ->setDescription('Sample 02 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // master slide
        $masterSlides = $objPHPPresentation->getAllMasterSlides();
        /** @var Slide $objMaster */
        $objMaster = reset($masterSlides);

        $objShape = $objMaster->createRichTextShape();
        $objShape->setWidthAndHeight(270, 30)->setOffsetX(600)->setOffsetY(655);
        $objShape->createTextRun("Footer")
            ->getFont()
            ->setName("Arial")
            ->setSize(7);
        $objShape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

        $objShape = $objMaster->createRichTextShape();
        $objShape->setWidthAndHeight(50, 30)->setOffsetX(870)->setOffsetY(655);
        $objShape->createTextRun("")
            ->getFont()
            ->setName("Arial")
            ->setSize(7);
        $objShape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_SLIDENUM));

        // slide with placeholder
        $currentSlide = $objPHPPresentation->getActiveSlide();

        $objShape = $currentSlide->createRichTextShape();
        $objShape->setWidthAndHeight(50, 30)->setOffsetX(870)->setOffsetY(655);
        $objShape->createTextRun("")
            ->getFont()
            ->setName("Arial")
            ->setSize(7);
        $objShape->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_SLIDENUM));

        return $objPHPPresentation;
    }
}
