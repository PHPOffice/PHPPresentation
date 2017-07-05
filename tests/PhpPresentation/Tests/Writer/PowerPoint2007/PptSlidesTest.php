<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\Common\Drawing;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Slide\Animation;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class PptSlideTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/42
     */
    public function testAlignmentShapeAuto()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_AUTO);
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeNotExists('ppt/slides/slide1.xml', $element, 'anchor');
    }

    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/42
     */
    public function testAlignmentShapeBase()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BASE);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeNotExists('ppt/slides/slide1.xml', $element, 'anchor');
    }

    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/35
     */
    public function testAlignmentShapeBottom()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BOTTOM);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'anchor', Alignment::VERTICAL_BOTTOM);
    }

    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/35
     */
    public function testAlignmentShapeCenter()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'anchor', Alignment::VERTICAL_CENTER);
    }

    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/35
     */
    public function testAlignmentShapeTop()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'anchor', Alignment::VERTICAL_TOP);
    }

    public function testAnimation()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape1 = $oSlide->createRichTextShape();
        $oShape2 = $oSlide->createRichTextShape();
        $oAnimation = new Animation();
        $oAnimation->addShape($oShape1);
        $oAnimation->addShape($oShape2);
        $oSlide->addAnimation($oAnimation);

        $element = '/p:sld/p:timing';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/p:sld/p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/p:sld/p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
    }

    public function testCommentRelationship()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->addShape(new Comment());

        $element = '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"]';
        $this->assertZipXmlElementExists('ppt/slides/_rels/slide1.xml.rels', $element);
    }

    public function testCommentInGroupRelationship()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oGroup = new Group();
        $oGroup->addShape(new Comment());
        $oSlide->addShape($oGroup);

        $element = '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"]';
        $this->assertZipXmlElementExists('ppt/slides/_rels/slide1.xml.rels', $element);
    }

    public function testDrawingWithHyperlink()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');

        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:nvPicPr/p:cNvPr/a:hlinkClick';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'r:id', 'rId3');
    }

    public function testDrawingShapeBorder()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        $oShape->getBorder()->setLineStyle(Border::LINE_DOUBLE);

        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:spPr/a:ln';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'cmpd', Border::LINE_DOUBLE);
    }

    public function testDrawingShapeShadow()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        $oShape->getShadow()->setVisible(true)->setDirection(45)->setDistance(10);

        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:spPr/a:effectLst/a:outerShdw';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
    }

    public function testFillGradientLinearTable()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FF' . $expected1))->setEndColor(new Color('FF' . $expected2));

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="0%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected1);
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="100%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected2);
    }

    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/61
     */
    public function testFillGradientLinearRichText()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oFill = $oShape->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FF' . $expected1))->setEndColor(new Color('FF' . $expected2));

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="0%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected1);
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="100%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected2);
    }

    public function testFillGradientPathTable()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_PATH)->setStartColor(new Color('FF' . $expected1))->setEndColor(new Color('FF' . $expected2));

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="0%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected1);
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="100%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected2);
    }

    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/61
     */
    public function testFillGradientPathText()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oFill = $oShape->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_PATH)->setStartColor(new Color('FF' . $expected1))->setEndColor(new Color('FF' . $expected2));

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="0%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected1);
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="100%"]/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected2);
    }

    public function testFillPatternTable()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_PATTERN_LIGHTTRELLIS)->setStartColor(new Color('FF' . $expected1))->setEndColor(new Color('FF' . $expected2));

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:pattFill/a:fgClr/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected1);
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:pattFill/a:bgClr/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected2);
    }

    public function testFillSolidTable()
    {
        $expected = 'E06B20';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF' . $expected));

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:solidFill/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected);
    }

    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/61
     */
    public function testFillSolidText()
    {
        $expected = 'E06B20';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oFill = $oShape->getFill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF' . $expected));

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:solidFill/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expected);
    }

    public function testHyperlink()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
    }

    public function testHyperlinkInternal()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setSlideNumber(1);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'action', 'ppaction://hlinksldjump');
    }

    public function testListBullet()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletColor(new Color('76543210'));

        $oExpectedFont = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletFont();
        $oExpectedChar = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletChar();
        $oExpectedColor = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletColor()->getRGB();
        $oExpectedAlpha = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletColor()->getAlpha() . "%";

        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:pPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:buFont');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element . '/a:buFont', 'typeface', $oExpectedFont);
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:buChar');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element . '/a:buChar', 'char', $oExpectedChar);
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:buClr');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:buClr/a:srgbClr');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element . '/a:buClr/a:srgbClr', 'val', $oExpectedColor);
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:buClr/a:srgbClr/a:alpha');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element . '/a:buClr/a:srgbClr/a:alpha', 'val', $oExpectedAlpha);
    }

    public function testListNumeric()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletNumericStyle(Bullet::NUMERIC_EA1CHSPERIOD);
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletNumericStartAt(5);
        $oExpectedFont = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletFont();
        $oExpectedChar = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletNumericStyle();
        $oExpectedStart = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletNumericStartAt();
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:pPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:buFont');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element . '/a:buFont', 'typeface', $oExpectedFont);
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:buAutoNum');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element . '/a:buAutoNum', 'type', $oExpectedChar);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element . '/a:buAutoNum', 'startAt', $oExpectedStart);
    }

    public function testLine()
    {
        $valEmu10 = Drawing::pixelsToEmu(10);
        $valEmu90 = Drawing::pixelsToEmu(90);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->createLineShape(10, 10, 100, 100);
        $oSlide->createLineShape(100, 10, 10, 100);
        $oSlide->createLineShape(10, 100, 100, 10);
        $oSlide->createLineShape(100, 100, 10, 10);

        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:prstGeom';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'prst', 'line');

        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:xfrm/a:ext[@cx="' . $valEmu90 . '"][@cy="' . $valEmu90 . '"]';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);

        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:xfrm/a:off[@x="' . $valEmu10 . '"][@y="' . $valEmu10 . '"]';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);

        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:xfrm[@flipV="1"]/a:off[@x="' . $valEmu10 . '"][@y="' . $valEmu10 . '"]';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
    }

    public function testMedia()
    {
        $expectedName = 'MyName';
        $expectedWidth = rand(1, 100);
        $expectedHeight = rand(1, 100);
        $expectedX = rand(1, 100);
        $expectedY = rand(1, 100);

        $oMedia = new Media();
        $oMedia->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/videos/sintel_trailer-480p.ogv')
            ->setName($expectedName)
            ->setResizeProportional(false)
            ->setHeight($expectedHeight)
            ->setWidth($expectedWidth)
            ->setOffsetX($expectedX)
            ->setOffsetY($expectedY);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->addShape($oMedia);

        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:nvPicPr/p:cNvPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'name', $expectedName);

        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:nvPicPr/p:nvPr/a:videoFile';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:nvPicPr/p:nvPr/p:extLst/p:ext';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'uri', '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}');
    }

    public function testNote()
    {
        $oLayout = $this->oPresentation->getLayout();
        $oSlide = $this->oPresentation->getActiveSlide();
        $oNote = $oSlide->getNote();
        $oRichText = $oNote->createRichTextShape()
            ->setHeight($oLayout->getCY($oLayout::UNIT_PIXEL))
            ->setWidth($oLayout->getCX($oLayout::UNIT_PIXEL))
            ->setOffsetX(170)
            ->setOffsetY(180);
        $oRichText->createTextRun('testNote');

        // Content Types
        // $element = '/Types/Override[@PartName="/ppt/notesSlides/notesSlide1.xml"][@ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"]';
        // $this->assertTrue($pres->elementExists($element, '[Content_Types].xml'));
        // Rels
        // $element = '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"][@Target="../notesSlides/notesSlide1.xml"]';
        // $this->assertTrue($pres->elementExists($element, 'ppt/slides/_rels/slide1.xml.rels'));
        // Slide

        $element = '/p:notes';
        $this->assertZipXmlElementExists('ppt/notesSlides/notesSlide1.xml', $element);
        // Slide Image Placeholder 1
        $element = '/p:notes/p:cSld/p:spTree/p:sp/p:nvSpPr/p:cNvPr[@id="2"][@name="Slide Image Placeholder 1"]';
        $this->assertZipXmlElementExists('ppt/notesSlides/notesSlide1.xml', $element);
        $element = '/p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:off';
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'x', 0);
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'y', 0);
        $element = '/p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:ext';
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'cx', Drawing::pixelsToEmu(round($oNote->getExtentX() / 2)));
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'cy', Drawing::pixelsToEmu(round($oNote->getExtentY() / 2)));

        // Notes Placeholder
        $element = '/p:notes/p:cSld/p:spTree/p:sp/p:nvSpPr/p:cNvPr[@id="3"][@name="Notes Placeholder"]';
        $this->assertZipXmlElementExists('ppt/notesSlides/notesSlide1.xml', $element);

        // Notes
        $element = '/p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm/a:off';
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'x', Drawing::pixelsToEmu($oNote->getOffsetX()));
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'y', Drawing::pixelsToEmu(round($oNote->getExtentY() / 2) + $oNote->getOffsetY()));
        $element = '/p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm/a:ext';
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'cx', 5486400);
        $this->assertZipXmlAttributeEquals('ppt/notesSlides/notesSlide1.xml', $element, 'cy', 3600450);
        $element = '/p:notes/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:t';
        $this->assertZipXmlElementExists('ppt/notesSlides/notesSlide1.xml', $element);
    }

    public function testRichTextAutoFitNormal()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->setAutoFit(RichText::AUTOFIT_NORMAL, 47.5, 20);
        $oRichText->createTextRun('This is my text for the test.');

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr/a:normAutofit';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'fontScale', 47500);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'lnSpcReduction', 20000);
    }

    public function testRichTextBreak()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createBreak();

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:br';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
    }

    public function testRichTextHyperlink()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getHyperLink()->setUrl('http://www.google.fr');

        $element = '/p:sld/p:cSld/p:spTree/p:sp//a:hlinkClick';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
    }

    public function testRichTextLineSpacing()
    {
        $expectedLineSpacing = rand(1, 100);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->getActiveParagraph()->setLineSpacing($expectedLineSpacing);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:pPr/a:lnSpc/a:spcPct';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'val', $expectedLineSpacing . '%');
    }

    public function testRichTextRunLanguage()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('MyText');

        $expectedElement = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('ppt/slides/slide1.xml', $expectedElement, 'lang');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $expectedElement, 'lang', 'en-US');

        $oRun->setLanguage('de_DE');
        $this->resetPresentationFile();

        $expectedElement = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('ppt/slides/slide1.xml', $expectedElement, 'lang');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $expectedElement, 'lang', 'de_DE');
    }

    public function testRichTextShadow()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->getShadow()->setVisible(true)->setAlpha(75)->setBlurRadius(2)->setDirection(45);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:effectLst/a:outerShdw';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
    }

    public function testRichTextUpright()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->setUpright(true);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'upright', '1');
    }

    public function testRichTextVertical()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->setVertical(true);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'vert', 'vert');
    }

    public function testSlideLayoutExists()
    {
        $this->assertZipFileExists('ppt/slideLayouts/slideLayout1.xml');
    }

    public function testStyleCharacterSpacing()
    {
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('pText');
        // Default : $oRun->getFont()->setCharacterSpacing(0);

        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'spc', '0');
        
        $oRun->getFont()->setCharacterSpacing(42);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'spc', '4200');
    }

    public function testStyleSubScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('pText');
        $oRun->getFont()->setSubScript(true);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'baseline', '-25000');
    }

    public function testStyleSuperScript()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('pText');
        $oRun->getFont()->setSuperScript(true);

        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'baseline', '30000');
    }

    public function testTableWithAlignment()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeNotExists('ppt/slides/slide1.xml', $element, 'anchor');
        $this->assertZipXmlAttributeNotExists('ppt/slides/slide1.xml', $element, 'vert');

        $oCell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BOTTOM);
        $oCell->getActiveParagraph()->getAlignment()->setTextDirection(Alignment::TEXT_DIRECTION_STACKED);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'anchor', Alignment::VERTICAL_BOTTOM);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'vert', Alignment::TEXT_DIRECTION_STACKED);
    }

    public function testTableWithBorder()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell(1);
        $oCell->createTextRun('AAA');
        $oCell->getBorders()->getBottom()->setDashStyle(Border::DASH_DASH);
        $oCell->getBorders()->getBottom()->setLineStyle(Border::LINE_DOUBLE);
        $oCell->getBorders()->getTop()->setDashStyle(Border::DASH_DASHDOT);
        $oCell->getBorders()->getTop()->setLineStyle(Border::LINE_SINGLE);
        $oCell->getBorders()->getRight()->setDashStyle(Border::DASH_DOT);
        $oCell->getBorders()->getRight()->setLineStyle(Border::LINE_THICKTHIN);
        $oCell->getBorders()->getLeft()->setDashStyle(Border::DASH_LARGEDASH);
        $oCell->getBorders()->getLeft()->setLineStyle(Border::LINE_THINTHICK);
        $oCell = $oRow->nextCell();
        $oCell->createTextRun('BBB');
        $oCell->getBorders()->getRight()->setDashStyle(Border::DASH_LARGEDASHDOT);
        $oCell->getBorders()->getRight()->setLineStyle(Border::LINE_TRI);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell(1);
        $oCell->createTextRun('CCC');
        $oCell->getBorders()->getBottom()->setDashStyle(Border::DASH_LARGEDASHDOTDOT);
        $oCell->getBorders()->getBottom()->setLineStyle(Border::LINE_NONE);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr';

        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnL[@cmpd="' . Border::LINE_THINTHICK . '"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnL[@cmpd="' . Border::LINE_THINTHICK . '"]/a:prstDash[@val="' . Border::DASH_LARGEDASH . '"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnR[@cmpd="' . Border::LINE_THICKTHIN . '"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnR[@cmpd="' . Border::LINE_THICKTHIN . '"]/a:prstDash[@val="' . Border::DASH_DOT . '"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnT[@cmpd="' . Border::LINE_SINGLE . '"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnT[@cmpd="' . Border::LINE_SINGLE . '"]/a:prstDash[@val="' . Border::DASH_DASHDOT . '"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnB[@cmpd="' . Border::LINE_SINGLE . '"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/a:lnB[@cmpd="' . Border::LINE_SINGLE . '"]/a:prstDash[@val="' . Border::DASH_SOLID . '"]');
    }

    public function testTableWithCellMargin()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->getActiveParagraph()->getAlignment()
            ->setMarginBottom(10)
            ->setMarginLeft(20)
            ->setMarginRight(30)
            ->setMarginTop(40);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'marB', Drawing::pixelsToEmu(10));
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'marL', Drawing::pixelsToEmu(20));
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'marR', Drawing::pixelsToEmu(30));
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'marT', Drawing::pixelsToEmu(40));
    }

    public function testTableWithColspan()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->setColSpan(2);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'gridSpan', 2);
    }

    public function testTableWithRowspan()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->setRowSpan(2);
        $oShape->createRow();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('BBB');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '[@rowSpan="2"]');
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '[@vMerge="1"]');
    }

    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/70
     */
    public function testTableWithHyperlink()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oTextRun = $oCell->createTextRun('AAA');
        $oHyperlink = $oTextRun->getHyperlink();
        $oHyperlink->setUrl('https://github.com/PHPOffice/PHPPresentation/');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'r:id', 'rId2');
    }

    public function testTransition()
    {
        $value = rand(1000, 5000);
        $element = '/p:sld/p:transition';

        $this->assertZipXmlElementNotExists('ppt/slides/slide1.xml', $element);

        $oTransition = new Transition();
        $oTransition->setTimeTrigger(true, $value);
        $this->oPresentation->getActiveSlide()->setTransition($oTransition);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $this->assertZipXmlAttributeExists('ppt/slides/slide1.xml', $element, 'advTm');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'advTm', $value);
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'advClick', '0');

        $oTransition->setSpeed(Transition::SPEED_FAST);
        $this->resetPresentationFile();
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'spd', 'fast');

        $oTransition->setSpeed(Transition::SPEED_MEDIUM);
        $this->resetPresentationFile();
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'spd', 'med');

        $oTransition->setSpeed(Transition::SPEED_SLOW);
        $this->resetPresentationFile();
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'spd', 'slow');

        $rcTransition = new \ReflectionClass('PhpOffice\PhpPresentation\Slide\Transition');
        $arrayConstants = $rcTransition->getConstants();
        foreach ($arrayConstants as $key => $value) {
            if (strpos($key, 'TRANSITION_') !== 0) {
                continue;
            }

            $oTransition->setTransitionType($rcTransition->getConstant($key));
            $this->oPresentation->getActiveSlide()->setTransition($oTransition);
            $this->resetPresentationFile();
            switch ($key) {
                case 'TRANSITION_BLINDS_HORIZONTAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:blinds[@dir=\'horz\']');
                    break;
                case 'TRANSITION_BLINDS_VERTICAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:blinds[@dir=\'vert\']');
                    break;
                case 'TRANSITION_CHECKER_HORIZONTAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:checker[@dir=\'horz\']');
                    break;
                case 'TRANSITION_CHECKER_VERTICAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:checker[@dir=\'vert\']');
                    break;
                case 'TRANSITION_CIRCLE_HORIZONTAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:circle[@dir=\'horz\']');
                    break;
                case 'TRANSITION_CIRCLE_VERTICAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:circle[@dir=\'vert\']');
                    break;
                case 'TRANSITION_COMB_HORIZONTAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:comb[@dir=\'horz\']');
                    break;
                case 'TRANSITION_COMB_VERTICAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:comb[@dir=\'vert\']');
                    break;
                case 'TRANSITION_COVER_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'d\']');
                    break;
                case 'TRANSITION_COVER_LEFT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'l\']');
                    break;
                case 'TRANSITION_COVER_LEFT_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'ld\']');
                    break;
                case 'TRANSITION_COVER_LEFT_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'lu\']');
                    break;
                case 'TRANSITION_COVER_RIGHT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'r\']');
                    break;
                case 'TRANSITION_COVER_RIGHT_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'rd\']');
                    break;
                case 'TRANSITION_COVER_RIGHT_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'ru\']');
                    break;
                case 'TRANSITION_COVER_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cover[@dir=\'u\']');
                    break;
                case 'TRANSITION_CUT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:cut');
                    break;
                case 'TRANSITION_DIAMOND':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:diamond');
                    break;
                case 'TRANSITION_DISSOLVE':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:dissolve');
                    break;
                case 'TRANSITION_FADE':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:fade');
                    break;
                case 'TRANSITION_NEWSFLASH':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:newsflash');
                    break;
                case 'TRANSITION_PLUS':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:plus');
                    break;
                case 'TRANSITION_PULL_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:pull[@dir=\'d\']');
                    break;
                case 'TRANSITION_PULL_LEFT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:pull[@dir=\'l\']');
                    break;
                case 'TRANSITION_PULL_RIGHT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:pull[@dir=\'r\']');
                    break;
                case 'TRANSITION_PULL_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:pull[@dir=\'u\']');
                    break;
                case 'TRANSITION_PUSH_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:push[@dir=\'d\']');
                    break;
                case 'TRANSITION_PUSH_LEFT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:push[@dir=\'l\']');
                    break;
                case 'TRANSITION_PUSH_RIGHT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:push[@dir=\'r\']');
                    break;
                case 'TRANSITION_PUSH_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:push[@dir=\'u\']');
                    break;
                case 'TRANSITION_RANDOM':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:random');
                    break;
                case 'TRANSITION_RANDOMBAR_HORIZONTAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:randomBar[@dir=\'horz\']');
                    break;
                case 'TRANSITION_RANDOMBAR_VERTICAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:randomBar[@dir=\'vert\']');
                    break;
                case 'TRANSITION_SPLIT_IN_HORIZONTAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:split[@dir=\'in\'][@orient=\'horz\']');
                    break;
                case 'TRANSITION_SPLIT_OUT_HORIZONTAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:split[@dir=\'out\'][@orient=\'horz\']');
                    break;
                case 'TRANSITION_SPLIT_IN_VERTICAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:split[@dir=\'in\'][@orient=\'vert\']');
                    break;
                case 'TRANSITION_SPLIT_OUT_VERTICAL':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:split[@dir=\'out\'][@orient=\'vert\']');
                    break;
                case 'TRANSITION_STRIPS_LEFT_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:strips[@dir=\'ld\']');
                    break;
                case 'TRANSITION_STRIPS_LEFT_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:strips[@dir=\'lu\']');
                    break;
                case 'TRANSITION_STRIPS_RIGHT_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:strips[@dir=\'rd\']');
                    break;
                case 'TRANSITION_STRIPS_RIGHT_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:strips[@dir=\'ru\']');
                    break;
                case 'TRANSITION_WEDGE':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:wedge');
                    break;
                case 'TRANSITION_WIPE_DOWN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:wipe[@dir=\'d\']');
                    break;
                case 'TRANSITION_WIPE_LEFT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:wipe[@dir=\'l\']');
                    break;
                case 'TRANSITION_WIPE_RIGHT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:wipe[@dir=\'r\']');
                    break;
                case 'TRANSITION_WIPE_UP':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:wipe[@dir=\'u\']');
                    break;
                case 'TRANSITION_ZOOM_IN':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:zoom[@dir=\'in\']');
                    break;
                case 'TRANSITION_ZOOM_OUT':
                    $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element . '/p:zoom[@dir=\'out\']');
                    break;
            }
        }

        $oTransition->setManualTrigger(true);
        $this->resetPresentationFile();
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $element, 'advClick', '1');
    }

    public function testVisibility()
    {
        $expectedElement = '/p:sld';

        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $expectedElement);
        $this->assertZipXmlAttributeNotExists('ppt/slides/slide1.xml', $expectedElement, 'show');

        $this->oPresentation->getActiveSlide()->setIsVisible(false);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('ppt/slides/slide1.xml', $expectedElement, 'show');
        $this->assertZipXmlAttributeEquals('ppt/slides/slide1.xml', $expectedElement, 'show', 0);
    }
}
