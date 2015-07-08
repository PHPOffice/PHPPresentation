<?php
/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\Common\Drawing;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\Slide;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class SlideTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }
    
    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Invalid \PhpOffice\PhpPresentation\Slide object passed.
     */
    public function testConstructException()
    {
        $oSlide = new Slide();
        $oSlide->writeSlide();
    }

    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/42
     */
    public function testAlignmentShapeAuto()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_AUTO);
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/42
     */
    public function testAlignmentShapeBase()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BASE);
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/35
     */
    public function testAlignmentShapeBottom()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BOTTOM);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_BOTTOM, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/35
     */
    public function testAlignmentShapeCenter()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_CENTER, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPresentation/issues/35
     */
    public function testAlignmentShapeTop()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_TOP, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    public function testDrawingWithHyperlink()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:nvPicPr/p:cNvPr/a:hlinkClick';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('rId3', $pres->getElementAttribute($element, 'r:id', 'ppt/slides/slide1.xml'));
    }
    
    public function testDrawingShapeBorder()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        $oShape->getBorder()->setLineStyle(Border::LINE_DOUBLE);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:spPr/a:ln';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Border::LINE_DOUBLE, $pres->getElementAttribute($element, 'cmpd', 'ppt/slides/slide1.xml'));
    }
    
    public function testDrawingShapeShadow()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        $oShape->getShadow()->setVisible(true)->setDirection(45)->setDistance(10);
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:spPr/a:effectLst/a:outerShdw';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testFillGradientLinearTable()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
        
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="0"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="100000"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/61
     */
    public function testFillGradientLinearRichText()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
    
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oFill = $oShape->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="0"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="100000"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testFillGradientPathTable()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
    
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_PATH)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="0"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="100000"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/61
     */
    public function testFillGradientPathText()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
        
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oFill = $oShape->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_PATH)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="0"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:gradFill/a:gsLst/a:gs[@pos="100000"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testFillPatternTable()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
        
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_PATTERN_LIGHTTRELLIS)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:pattFill/a:fgClr/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:pattFill/a:bgClr/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testFillSolidTable()
    {
        $expected = 'E06B20';
    
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF'.$expected));
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:solidFill/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }

    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/61
     */
    public function testFillSolidText()
    {
        $expected = 'E06B20';
    
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oFill = $oShape->getFill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF'.$expected));
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:solidFill/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testHyperlink()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testHyperlinkInternal()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setSlideNumber(1);
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('ppaction://hlinksldjump', $pres->getElementAttribute($element, 'action', 'ppt/slides/slide1.xml'));
    }

    public function testListBullet()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oExpectedFont = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletFont();
        $oExpectedChar = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletChar();
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:pPr';
        $this->assertTrue($pres->elementExists($element.'/a:buFont', 'ppt/slides/slide1.xml'));
        $this->assertEquals($oExpectedFont, $pres->getElementAttribute($element.'/a:buFont', 'typeface', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:buChar', 'ppt/slides/slide1.xml'));
        $this->assertEquals($oExpectedChar, $pres->getElementAttribute($element.'/a:buChar', 'char', 'ppt/slides/slide1.xml'));
    }

    public function testListNumeric()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
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
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:pPr';
        $this->assertTrue($pres->elementExists($element.'/a:buFont', 'ppt/slides/slide1.xml'));
        $this->assertEquals($oExpectedFont, $pres->getElementAttribute($element.'/a:buFont', 'typeface', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:buAutoNum', 'ppt/slides/slide1.xml'));
        $this->assertEquals($oExpectedChar, $pres->getElementAttribute($element.'/a:buAutoNum', 'type', 'ppt/slides/slide1.xml'));
        $this->assertEquals($oExpectedStart, $pres->getElementAttribute($element.'/a:buAutoNum', 'startAt', 'ppt/slides/slide1.xml'));
    }
    
    public function testLine()
    {
        $valEmu10 = Drawing::pixelsToEmu(10);
        $valEmu90 = Drawing::pixelsToEmu(90);
        
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oSlide->createLineShape(10, 10, 100, 100);
        $oSlide->createLineShape(100, 10, 10, 100);
        $oSlide->createLineShape(10, 100, 100, 10);
        $oSlide->createLineShape(100, 100, 10, 10);
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:prstGeom';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('line', $pres->getElementAttribute($element, 'prst', 'ppt/slides/slide1.xml'));
        
        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:xfrm/a:ext[@cx="'.$valEmu90.'"][@cy="'.$valEmu90.'"]';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        
        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:xfrm/a:off[@x="'.$valEmu10.'"][@y="'.$valEmu10.'"]';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        
        $element = '/p:sld/p:cSld/p:spTree/p:cxnSp/p:spPr/a:xfrm[@flipV="1"]/a:off[@x="'.$valEmu10.'"][@y="'.$valEmu10.'"]';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }

    public function testNote()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oNote = $oSlide->getNote();
        $oRichText = $oNote->createRichTextShape()->setHeight(300)->setWidth(600);
        $oRichText->createTextRun('testNote');
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        // Content Types
        // $element = '/Types/Override[@PartName="/ppt/notesSlides/notesSlide1.xml"][@ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"]';
        // $this->assertTrue($pres->elementExists($element, '[Content_Types].xml'));
        // Rels
        // $element = '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"][@Target="../notesSlides/notesSlide1.xml"]';
        // $this->assertTrue($pres->elementExists($element, 'ppt/slides/_rels/slide1.xml.rels'));
        // Slide
        $element = '/p:notes';
        $this->assertTrue($pres->elementExists($element, 'ppt/notesSlides/notesSlide1.xml'));
        $element = '/p:notes/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:t';
        $this->assertTrue($pres->elementExists($element, 'ppt/notesSlides/notesSlide1.xml'));
    }
    
    public function testRichTextAutoFitNormal()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->setAutoFit(RichText::AUTOFIT_NORMAL, 47.5, 20);
        $oRichText->createTextRun('This is my text for the test.');
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr/a:normAutofit';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(47500, $pres->getElementAttribute($element, 'fontScale', 'ppt/slides/slide1.xml'));
        $this->assertEquals(20000, $pres->getElementAttribute($element, 'lnSpcReduction', 'ppt/slides/slide1.xml'));
    }
    
    public function testRichTextBreak()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createBreak();
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:br';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testRichTextHyperlink()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getHyperLink()->setUrl('http://www.google.fr');

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');

        $element = '/p:sld/p:cSld/p:spTree/p:sp//a:hlinkClick';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testRichTextShadow()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->getShadow()->setVisible(true)->setAlpha(75)->setBlurRadius(2)->setDirection(45);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:spPr/a:effectLst/a:outerShdw';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testRichTextUpright()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->setUpright(true);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('1', $pres->getElementAttribute($element, 'upright', 'ppt/slides/slide1.xml'));
    }
    
    public function testRichTextVertical()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->setVertical(true);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('vert', $pres->getElementAttribute($element, 'vert', 'ppt/slides/slide1.xml'));
    }

    public function testStyleSubScript()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('pText');
        $oRun->getFont()->setSubScript(true);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('-25000', $pres->getElementAttribute($element, 'baseline', 'ppt/slides/slide1.xml'));
    }

    public function testStyleSuperScript()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('pText');
        $oRun->getFont()->setSuperScript(true);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('30000', $pres->getElementAttribute($element, 'baseline', 'ppt/slides/slide1.xml'));
    }

    public function testTableWithAlignment()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BOTTOM);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_BOTTOM, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }

    public function testTableWithBorder()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
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
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr';
        $this->assertTrue($pres->elementExists($element.'/a:lnL[@cmpd="'.Border::LINE_THINTHICK.'"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:lnL[@cmpd="'.Border::LINE_THINTHICK.'"]/a:prstDash[@val="'.Border::DASH_LARGEDASH.'"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:lnR[@cmpd="'.Border::LINE_THICKTHIN.'"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:lnR[@cmpd="'.Border::LINE_THICKTHIN.'"]/a:prstDash[@val="'.Border::DASH_DOT.'"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:lnT[@cmpd="'.Border::LINE_SINGLE.'"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:lnT[@cmpd="'.Border::LINE_SINGLE.'"]/a:prstDash[@val="'.Border::DASH_DASHDOT.'"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:lnB[@cmpd="'.Border::LINE_SINGLE.'"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:lnB[@cmpd="'.Border::LINE_SINGLE.'"]/a:prstDash[@val="'.Border::DASH_SOLID.'"]', 'ppt/slides/slide1.xml'));
    }

    public function testTableWithColspan()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->setColSpan(2);
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(2, $pres->getElementAttribute($element, 'gridSpan', 'ppt/slides/slide1.xml'));
    }

    public function testTableWithRowspan()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->setRowSpan(2);
        $oRow = $oShape->createRow();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('BBB');
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc';
        $this->assertTrue($pres->elementExists($element.'[@rowSpan="2"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'[@vMerge="1"]', 'ppt/slides/slide1.xml'));
    }

    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/70
     */
    public function testTableWithHyperlink()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oTextRun = $oCell->createTextRun('AAA');
        $oHyperlink = $oTextRun->getHyperlink();
        $oHyperlink->setUrl('https://github.com/PHPOffice/PHPPresentation/');
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('rId2', $pres->getElementAttribute($element, 'r:id', 'ppt/slides/slide1.xml'));
    }

    public function testTransition()
    {
        $value = rand(1000, 5000);

        $oTransition = new Transition();

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();

        $element = '/p:sld/p:transition';

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->elementExists($element, 'ppt/slides/slide1.xml'));

        $oTransition->setTimeTrigger(true, $value);
        $oSlide->setTransition($oTransition);

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'advTm', 'ppt/slides/slide1.xml'));
        $this->assertEquals($value, $pres->getElementAttribute($element, 'advTm', 'ppt/slides/slide1.xml'));
        $this->assertContains('0', $pres->getElementAttribute($element, 'advClick', 'ppt/slides/slide1.xml'));

        $oTransition->setSpeed(Transition::SPEED_FAST);
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertContains('fast', $pres->getElementAttribute($element, 'spd', 'ppt/slides/slide1.xml'));

        $oTransition->setSpeed(Transition::SPEED_MEDIUM);
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertContains('med', $pres->getElementAttribute($element, 'spd', 'ppt/slides/slide1.xml'));

        $oTransition->setSpeed(Transition::SPEED_SLOW);
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertContains('slow', $pres->getElementAttribute($element, 'spd', 'ppt/slides/slide1.xml'));

        $rcTransition = new \ReflectionClass('PhpOffice\PhpPresentation\Slide\Transition');
        $arrayConstants = $rcTransition->getConstants();
        foreach ($arrayConstants as $key => $value) {
            if (strpos($key, 'TRANSITION_') !== 0) {
                continue;
            }

            $oTransition->setTransitionType($rcTransition->getConstant($key));
            $oSlide->setTransition($oTransition);
            $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
            switch ($key) {
                case 'TRANSITION_BLINDS_HORIZONTAL':
                    $this->assertTrue($pres->elementExists($element.'/p:blinds[@dir=\'horz\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_BLINDS_VERTICAL':
                    $this->assertTrue($pres->elementExists($element.'/p:blinds[@dir=\'vert\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_CHECKER_HORIZONTAL':
                    $this->assertTrue($pres->elementExists($element.'/p:checker[@dir=\'horz\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_CHECKER_VERTICAL':
                    $this->assertTrue($pres->elementExists($element.'/p:checker[@dir=\'vert\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_CIRCLE_HORIZONTAL':
                    $this->assertTrue($pres->elementExists($element.'/p:circle[@dir=\'horz\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_CIRCLE_VERTICAL':
                    $this->assertTrue($pres->elementExists($element.'/p:circle[@dir=\'vert\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COMB_HORIZONTAL':
                    $this->assertTrue($pres->elementExists($element.'/p:comb[@dir=\'horz\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COMB_VERTICAL':
                    $this->assertTrue($pres->elementExists($element.'/p:comb[@dir=\'vert\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'d\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_LEFT':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'l\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_LEFT_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'ld\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_LEFT_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'lu\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_RIGHT':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'r\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_RIGHT_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'rd\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_RIGHT_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'ru\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_COVER_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:cover[@dir=\'u\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_CUT':
                    $this->assertTrue($pres->elementExists($element.'/p:cut', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_DIAMOND':
                    $this->assertTrue($pres->elementExists($element.'/p:diamond', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_DISSOLVE':
                    $this->assertTrue($pres->elementExists($element.'/p:dissolve', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_FADE':
                    $this->assertTrue($pres->elementExists($element.'/p:fade', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_NEWSFLASH':
                    $this->assertTrue($pres->elementExists($element.'/p:newsflash', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PLUS':
                    $this->assertTrue($pres->elementExists($element.'/p:plus', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PULL_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:pull[@dir=\'d\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PULL_LEFT':
                    $this->assertTrue($pres->elementExists($element.'/p:pull[@dir=\'l\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PULL_RIGHT':
                    $this->assertTrue($pres->elementExists($element.'/p:pull[@dir=\'r\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PULL_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:pull[@dir=\'u\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PUSH_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:push[@dir=\'d\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PUSH_LEFT':
                    $this->assertTrue($pres->elementExists($element.'/p:push[@dir=\'l\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PUSH_RIGHT':
                    $this->assertTrue($pres->elementExists($element.'/p:push[@dir=\'r\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_PUSH_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:push[@dir=\'u\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_RANDOM':
                    $this->assertTrue($pres->elementExists($element.'/p:random', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_RANDOMBAR_HORIZONTAL':
                    $this->assertTrue($pres->elementExists($element.'/p:randomBar[@dir=\'horz\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_RANDOMBAR_VERTICAL':
                    $this->assertTrue($pres->elementExists($element.'/p:randomBar[@dir=\'vert\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_SPLIT_IN_HORIZONTAL':
                    $this->assertTrue($pres->elementExists($element.'/p:split[@dir=\'in\'][@orient=\'horz\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_SPLIT_OUT_HORIZONTAL':
                    $this->assertTrue($pres->elementExists($element.'/p:split[@dir=\'out\'][@orient=\'horz\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_SPLIT_IN_VERTICAL':
                    $this->assertTrue($pres->elementExists($element.'/p:split[@dir=\'in\'][@orient=\'vert\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_SPLIT_OUT_VERTICAL':
                    $this->assertTrue($pres->elementExists($element.'/p:split[@dir=\'out\'][@orient=\'vert\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_STRIPS_LEFT_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:strips[@dir=\'ld\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_STRIPS_LEFT_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:strips[@dir=\'lu\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_STRIPS_RIGHT_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:strips[@dir=\'rd\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_STRIPS_RIGHT_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:strips[@dir=\'ru\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_WEDGE':
                    $this->assertTrue($pres->elementExists($element.'/p:wedge', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_WIPE_DOWN':
                    $this->assertTrue($pres->elementExists($element.'/p:wipe[@dir=\'d\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_WIPE_LEFT':
                    $this->assertTrue($pres->elementExists($element.'/p:wipe[@dir=\'l\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_WIPE_RIGHT':
                    $this->assertTrue($pres->elementExists($element.'/p:wipe[@dir=\'r\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_WIPE_UP':
                    $this->assertTrue($pres->elementExists($element.'/p:wipe[@dir=\'u\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_ZOOM_IN':
                    $this->assertTrue($pres->elementExists($element.'/p:zoom[@dir=\'in\']', 'ppt/slides/slide1.xml'));
                    break;
                case 'TRANSITION_ZOOM_OUT':
                    $this->assertTrue($pres->elementExists($element.'/p:zoom[@dir=\'out\']', 'ppt/slides/slide1.xml'));
                    break;
            }
        }

        $oTransition->setManualTrigger(true);
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertContains('1', $pres->getElementAttribute($element, 'advClick', 'ppt/slides/slide1.xml'));
    }
}
