<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shared\Drawing;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Bullet;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Tests\TestHelperDOCX;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Shape\Hyperlink;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Slide;

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
     * @expectedExceptionMessage Invalid \PhpOffice\PhpPowerpoint\Slide object passed.
     */
    public function testConstructException()
    {
        $oSlide = new Slide();
        $oSlide->writeSlide();
    }

    /**
     * @link https://github.com/PHPOffice/PHPPowerPoint/issues/42
     */
    public function testAlignmentShapeAuto()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_AUTO);
    
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPowerPoint/issues/42
     */
    public function testAlignmentShapeBase()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BASE);
    
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPowerPoint/issues/35
     */
    public function testAlignmentShapeBottom()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BOTTOM);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_BOTTOM, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPowerPoint/issues/35
     */
    public function testAlignmentShapeCenter()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_CENTER, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    /**
     * @link https://github.com/PHPOffice/PHPPowerPoint/issues/35
     */
    public function testAlignmentShapeTop()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createRichTextShape()->setWidth(400)->setHeight(400)->setOffsetX(100)->setOffsetY(100);
        $oShape->createTextRun('this text should be vertically aligned');
        $oShape->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_TOP);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_TOP, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }
    
    public function testDrawingShapeBorder()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPOWERPOINT_TESTS_BASE_DIR.'/resources/images/PHPPowerPointLogo.png');
        $oShape->getBorder()->setLineStyle(Border::LINE_DOUBLE);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:spPr/a:ln';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Border::LINE_DOUBLE, $pres->getElementAttribute($element, 'cmpd', 'ppt/slides/slide1.xml'));
    }
    
    public function testDrawingShapeShadow()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPOWERPOINT_TESTS_BASE_DIR.'/resources/images/PHPPowerPointLogo.png');
        $oShape->getShadow()->setVisible(true)->setDirection(45)->setDistance(10);
    
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:pic/p:spPr/a:effectLst/a:outerShdw';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testFillGradientLinear()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
        
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="0"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="100000"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testFillGradientPath()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
        
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_GRADIENT_PATH)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="0"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:gradFill/a:gsLst/a:gs[@pos="100000"]/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testFillPattern()
    {
        $expected1 = 'E06B20';
        $expected2 = strrev($expected1);
        
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_PATTERN_LIGHTTRELLIS)->setStartColor(new Color('FF'.$expected1))->setEndColor(new Color('FF'.$expected2));
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:pattFill/a:fgClr/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected1, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:pattFill/a:bgClr/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected2, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testFillSolid()
    {
        $expected = 'E06B20';
    
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape(1);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('R1C1');
        $oFill = $oCell->getFill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF'.$expected));
    
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr/a:solidFill/a:srgbClr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals($expected, $pres->getElementAttribute($element, 'val', 'ppt/slides/slide1.xml'));
    }
    
    public function testHyperlink()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
    
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testHyperlinkInternal()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setSlideNumber(1);
    
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
    
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr/a:hlinkClick';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('ppaction://hlinksldjump', $pres->getElementAttribute($element, 'action', 'ppt/slides/slide1.xml'));
    }

    public function testListBullet()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oExpectedFont = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletFont();
        $oExpectedChar = $oRichText->getActiveParagraph()->getBulletStyle()->getBulletChar();
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:pPr';
        $this->assertTrue($pres->elementExists($element.'/a:buFont', 'ppt/slides/slide1.xml'));
        $this->assertEquals($oExpectedFont, $pres->getElementAttribute($element.'/a:buFont', 'typeface', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'/a:buChar', 'ppt/slides/slide1.xml'));
        $this->assertEquals($oExpectedChar, $pres->getElementAttribute($element.'/a:buChar', 'char', 'ppt/slides/slide1.xml'));
    }

    public function testListNumeric()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
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
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
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
        
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oSlide->createLineShape(10, 10, 100, 100);
        $oSlide->createLineShape(100, 10, 10, 100);
        $oSlide->createLineShape(10, 100, 100, 10);
        $oSlide->createLineShape(100, 100, 10, 10);
    
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
    
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

    public function testRichTextBreak()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createBreak();
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:br';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
    }
    
    public function testRichTextUpright()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->setUpright(true);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('1', $pres->getElementAttribute($element, 'upright', 'ppt/slides/slide1.xml'));
    }
    
    public function testRichTextVertical()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->setVertical(true);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:bodyPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('vert', $pres->getElementAttribute($element, 'vert', 'ppt/slides/slide1.xml'));
    }

    public function testStyleSubScript()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('pText');
        $oRun->getFont()->setSubScript(true);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('-25000', $pres->getElementAttribute($element, 'baseline', 'ppt/slides/slide1.xml'));
    }

    public function testStyleSuperScript()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('pText');
        $oRun->getFont()->setSuperScript(true);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:sp/p:txBody/a:p/a:r/a:rPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals('30000', $pres->getElementAttribute($element, 'baseline', 'ppt/slides/slide1.xml'));
    }

    public function testTableWithAlignment()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->getActiveParagraph()->getAlignment()->setVertical(Alignment::VERTICAL_BOTTOM);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:tcPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(Alignment::VERTICAL_BOTTOM, $pres->getElementAttribute($element, 'anchor', 'ppt/slides/slide1.xml'));
    }

    public function testTableWithBorder()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
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
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
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
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->createTextRun('AAA');
        $oCell->setColSpan(2);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc';
        $this->assertTrue($pres->elementExists($element, 'ppt/slides/slide1.xml'));
        $this->assertEquals(2, $pres->getElementAttribute($element, 'gridSpan', 'ppt/slides/slide1.xml'));
    }

    public function testTableWithRowspan()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
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
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'PowerPoint2007');
        
        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc';
        $this->assertTrue($pres->elementExists($element.'[@rowSpan="2"]', 'ppt/slides/slide1.xml'));
        $this->assertTrue($pres->elementExists($element.'[@vMerge="1"]', 'ppt/slides/slide1.xml'));
    }
}
