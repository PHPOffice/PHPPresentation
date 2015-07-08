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

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Writer\ODPresentation;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;
use PhpOffice\Common\Drawing;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\PhpOffice\PhpPresentation\Style;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 */
class ContentTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }

    public function testDrawingWithHyperlink()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');
    
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
    
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/office:event-listeners/presentation:event-listener';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('https://github.com/PHPOffice/PHPPresentation/', $pres->getElementAttribute($element, 'xlink:href', 'content.xml'));
    }

    public function testGroup()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oShapeGroup = $oSlide->createGroup();
        $oShape = $oShapeGroup->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');
    
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
    
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:g';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:g/draw:frame/office:event-listeners/presentation:event-listener';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }
    
    public function testList()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testInnerList()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        
        $oRichText = $oSlide->createRichTextShape();
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
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:list/text:list-item/text:p/text:span';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testParagraphRichText()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('Alpha');
        $oRichText->createBreak();
        $oRichText->createText('Beta');
        $oRichText->createBreak();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:line-break';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:a';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('http://www.google.fr', $pres->getElementAttribute($element, 'xlink:href', 'content.xml'));
    }

    public function testListWithRichText()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRun = $oRichText->createTextRun('Alpha');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
        $oRichText->createBreak();
        $oRichText->createTextRun('Beta');
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:a';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:line-break';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testNote()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oNote = $oSlide->getNote();
        $oRichText = $oNote->createRichTextShape()->setHeight(300)->setWidth(600);
        $oRichText->createTextRun('testNote');
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:body/office:presentation/draw:page/presentation:notes';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $element = '/office:document-content/office:body/office:presentation/draw:page/presentation:notes/draw:frame/draw:text-box/text:p/text:span';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testRichTextAutoShrink()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertFalse($pres->attributeElementExists($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'draw:auto-grow-width', 'content.xml'));

        $oRichText1->setAutoShrinkHorizontal(false);
        $oRichText1->setAutoShrinkVertical(true);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-width', 'content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'draw:auto-grow-width', 'content.xml'));
        

        $oRichText1->setAutoShrinkHorizontal(true);
        $oRichText1->setAutoShrinkVertical(false);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-width', 'content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'draw:auto-grow-width', 'content.xml'));
    }

    public function testRichTextBorder()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';

        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_NONE);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'svg:stroke-color', 'content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'svg:stroke-width', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:stroke', 'content.xml'));
        $this->assertEquals('none', $pres->getElementAttribute($element, 'draw:stroke', 'content.xml'));
        
        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_SINGLE);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'svg:stroke-color', 'content.xml'));
        $this->assertEquals('#'.$oRichText1->getBorder()->getColor()->getRGB(), $pres->getElementAttribute($element, 'svg:stroke-color', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'svg:stroke-width', 'content.xml'));
        $this->assertStringEndsWith('cm', $pres->getElementAttribute($element, 'svg:stroke-width', 'content.xml'));
        $this->assertStringStartsWith((string) number_format(CommonDrawing::pointsToCentimeters($oRichText1->getBorder()->getLineWidth()), 3, '.', ''), $pres->getElementAttribute($element, 'svg:stroke-width', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:stroke', 'content.xml'));
        $this->assertEquals('solid', $pres->getElementAttribute($element, 'draw:stroke', 'content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'draw:stroke-dash', 'content.xml'));

        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_DASH);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertEquals('dash', $pres->getElementAttribute($element, 'draw:stroke', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:stroke-dash', 'content.xml'));
        $this->assertStringStartsWith('strokeDash_', $pres->getElementAttribute($element, 'draw:stroke-dash', 'content.xml'));
        $this->assertStringEndsWith($oRichText1->getBorder()->getDashStyle(), $pres->getElementAttribute($element, 'draw:stroke-dash', 'content.xml'));
    }
    
    public function testRichTextShadow()
    {
        $randAlpha = rand(0, 100);
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->getShadow()->setVisible(true)->setAlpha($randAlpha)->setBlurRadius(2);
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        for ($inc = 0; $inc <= 360; $inc += 45) {
            $randDistance = rand(0, 100);
            $oRichText->getShadow()->setDirection($inc)->setDistance($randDistance);
            $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
            $this->assertTrue($pres->elementExists($element, 'content.xml'));
            $this->assertEquals('visible', $pres->getElementAttribute($element, 'draw:shadow', 'content.xml'));
            $this->assertEquals('none', $pres->getElementAttribute($element, 'style:mirror', 'content.xml'));
            // Opacity
            $this->assertStringStartsWith((string)(100 - $randAlpha), $pres->getElementAttribute($element, 'draw:shadow-opacity', 'content.xml'));
            $this->assertStringEndsWith('%', $pres->getElementAttribute($element, 'draw:shadow-opacity', 'content.xml'));
            // Color
            $this->assertStringStartsWith('#', $pres->getElementAttribute($element, 'draw:shadow-color', 'content.xml'));
            // X
            $xOffset = $pres->getElementAttribute($element, 'draw:shadow-offset-x', 'content.xml');
            if ($inc == 90 || $inc == 270) {
                $this->assertEquals('0cm', $xOffset);
            } else {
                if ($inc > 90 && $inc < 270) {
                    $this->assertEquals('-'.Drawing::pixelsToCentimeters($randDistance).'cm', $xOffset);
                } else {
                    $this->assertEquals(Drawing::pixelsToCentimeters($randDistance).'cm', $xOffset);
                }
            }
            // Y
            $yOffset = $pres->getElementAttribute($element, 'draw:shadow-offset-y', 'content.xml');
            if ($inc == 0 || $inc == 180 || $inc == 360) {
                $this->assertEquals('0cm', $yOffset);
            } else {
                if (($inc > 0 && $inc < 180) || $inc == 360) {
                    $this->assertEquals(Drawing::pixelsToCentimeters($randDistance).'cm', $yOffset);
                } else {
                    $this->assertEquals('-'.Drawing::pixelsToCentimeters($randDistance).'cm', $yOffset);
                }
            }
        }
    }
    
    public function testStyleAlignment()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        $oRichText1->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $oRichText1->createTextRun('Run1');
        $p1HashCode = $oRichText1->getActiveParagraph()->getHashCode();
        
        $oRichText2 = $oSlide->createRichTextShape();
        $oRichText2->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_DISTRIBUTED);
        $oRichText2->createTextRun('Run2');
        $p2HashCode = $oRichText2->getActiveParagraph()->getHashCode();
        
        $oRichText3 = $oSlide->createRichTextShape();
        $oRichText3->getActiveParagraph()->getAlignment()->setHorizontal('AAAAA');
        $oRichText3->createTextRun('Run3');
        $p3HashCode = $oRichText3->getActiveParagraph()->getHashCode();
        
        $oRichText4 = $oSlide->createRichTextShape();
        $oRichText4->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);
        $oRichText4->createTextRun('Run4');
        $p4HashCode = $oRichText4->getActiveParagraph()->getHashCode();
        
        $oRichText5 = $oSlide->createRichTextShape();
        $oRichText5->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $oRichText5->createTextRun('Run5');
        $p5HashCode = $oRichText5->getActiveParagraph()->getHashCode();
        
        $oRichText6 = $oSlide->createRichTextShape();
        $oRichText6->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $oRichText6->createTextRun('Run6');
        $p6HashCode = $oRichText6->getActiveParagraph()->getHashCode();
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p1HashCode.'\']/style:paragraph-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('center', $pres->getElementAttribute($element, 'fo:text-align', 'content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p2HashCode.'\']/style:paragraph-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('justify', $pres->getElementAttribute($element, 'fo:text-align', 'content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p3HashCode.'\']/style:paragraph-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('left', $pres->getElementAttribute($element, 'fo:text-align', 'content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p4HashCode.'\']/style:paragraph-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('justify', $pres->getElementAttribute($element, 'fo:text-align', 'content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p5HashCode.'\']/style:paragraph-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('left', $pres->getElementAttribute($element, 'fo:text-align', 'content.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'P_'.$p6HashCode.'\']/style:paragraph-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('right', $pres->getElementAttribute($element, 'fo:text-align', 'content.xml'));
    }
    
    public function testStyleFont()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setBold(true);
        
        $expectedHashCode = $oRun->getFont()->getHashCode();
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'T_'.$expectedHashCode.'\']/style:text-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('bold', $pres->getElementAttribute($element, 'fo:font-weight', 'content.xml'));
    }
    
    public function testTable()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oSlide->createTableShape();
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }
    
    public function testTableCellFill()
    {
        $oColor = new Color();
        $oColor->setRGB(Color::COLOR_BLUE);
        
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor($oColor);
        
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->setFill($oFill);
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1r0c0\']';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('table-cell', $pres->getElementAttribute($element, 'style:family', 'content.xml'));
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1r0c0\']/style:graphic-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('solid', $pres->getElementAttribute($element, 'draw:fill', 'content.xml'));
        $this->assertStringStartsWith('#', $pres->getElementAttribute($element, 'draw:fill-color', 'content.xml'));
        $this->assertStringEndsWith($oColor->getRGB(), $pres->getElementAttribute($element, 'draw:fill-color', 'content.xml'));
    }
    
    public function testTableWithColspan()
    {
        $value = rand(2, 100);
        
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape($value);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->setColSpan($value);
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals($value, $pres->getElementAttribute($element, 'table:number-columns-spanned', 'content.xml'));
    }
    
    /**
     * @link : https://github.com/PHPOffice/PHPPresentation/issues/70
     */
    public function testTableWithHyperlink()
    {
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oTextRun = $oCell->createTextRun('AAA');
        $oHyperlink = $oTextRun->getHyperlink();
        $oHyperlink->setUrl('https://github.com/PHPOffice/PHPPresentation/');
    
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
    
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span/text:a';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('https://github.com/PHPOffice/PHPPresentation/', $pres->getElementAttribute($element, 'xlink:href', 'content.xml'));
    }
    
    public function testTableWithText()
    {
        $oRun = new Run();
        $oRun->setText('Test');
        
        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->addText($oRun);
        $oCell->createBreak();
        
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('Test', $pres->getElement($element, 'content.xml')->nodeValue);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span/text:line-break';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testTransition()
    {
        $value = rand(1000, 5000);

        $oTransition = new Transition();

        $phpPresentation = new PhpPresentation();
        $oSlide = $phpPresentation->getActiveSlide();

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePage0\']/style:drawing-page-properties';

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'presentation:duration', 'content.xml'));

        $oTransition->setTimeTrigger(true, $value);
        $oSlide->setTransition($oTransition);

        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'presentation:duration', 'content.xml'));
        $this->assertStringStartsWith('PT', $pres->getElementAttribute($element, 'presentation:duration', 'content.xml'));
        $this->assertStringEndsWith('S', $pres->getElementAttribute($element, 'presentation:duration', 'content.xml'));
        $this->assertContains(number_format($value / 1000, 6, '.', ''), $pres->getElementAttribute($element, 'presentation:duration', 'content.xml'));
        $this->assertContains('automatic', $pres->getElementAttribute($element, 'presentation:transition-type', 'content.xml'));

        $oTransition->setSpeed(Transition::SPEED_FAST);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertContains('fast', $pres->getElementAttribute($element, 'presentation:transition-speed', 'content.xml'));

        $oTransition->setSpeed(Transition::SPEED_MEDIUM);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertContains('medium', $pres->getElementAttribute($element, 'presentation:transition-speed', 'content.xml'));

        $oTransition->setSpeed(Transition::SPEED_SLOW);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertContains('slow', $pres->getElementAttribute($element, 'presentation:transition-speed', 'content.xml'));

        $rcTransition = new \ReflectionClass('PhpOffice\PhpPresentation\Slide\Transition');
        $arrayConstants = $rcTransition->getConstants();
        foreach ($arrayConstants as $key => $value) {
            if (strpos($key, 'TRANSITION_') !== 0) {
                continue;
            }

            $oTransition->setTransitionType($rcTransition->getConstant($key));
            $oSlide->setTransition($oTransition);
            $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
            switch ($key) {
                case 'TRANSITION_BLINDS_HORIZONTAL':
                    $this->assertContains('horizontal-stripes', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_BLINDS_VERTICAL':
                    $this->assertContains('vertical-stripes', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_CHECKER_HORIZONTAL':
                    $this->assertContains('horizontal-checkerboard', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_CHECKER_VERTICAL':
                    $this->assertContains('vertical-checkerboard', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_CIRCLE_HORIZONTAL':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_CIRCLE_VERTICAL':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COMB_HORIZONTAL':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COMB_VERTICAL':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_DOWN':
                    $this->assertContains('uncover-to-bottom', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_LEFT':
                    $this->assertContains('uncover-to-left', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_LEFT_DOWN':
                    $this->assertContains('uncover-to-lowerleft', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_LEFT_UP':
                    $this->assertContains('uncover-to-upperleft', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_RIGHT':
                    $this->assertContains('uncover-to-right', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_RIGHT_DOWN':
                    $this->assertContains('uncover-to-lowerright', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_RIGHT_UP':
                    $this->assertContains('uncover-to-upperright', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_COVER_UP':
                    $this->assertContains('uncover-to-top', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_CUT':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_DIAMOND':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_DISSOLVE':
                    $this->assertContains('dissolve', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_FADE':
                    $this->assertContains('fade-from-center', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_NEWSFLASH':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PLUS':
                    $this->assertContains('close', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PULL_DOWN':
                    $this->assertContains('stretch-from-bottom', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PULL_LEFT':
                    $this->assertContains('stretch-from-left', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PULL_RIGHT':
                    $this->assertContains('stretch-from-right', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PULL_UP':
                    $this->assertContains('stretch-from-top', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PUSH_DOWN':
                    $this->assertContains('roll-from-bottom', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PUSH_LEFT':
                    $this->assertContains('roll-from-left', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PUSH_RIGHT':
                    $this->assertContains('roll-from-right', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_PUSH_UP':
                    $this->assertContains('roll-from-top', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_RANDOM':
                    $this->assertContains('random', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_RANDOMBAR_HORIZONTAL':
                    $this->assertContains('horizontal-lines', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_RANDOMBAR_VERTICAL':
                    $this->assertContains('vertical-lines', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_SPLIT_IN_HORIZONTAL':
                    $this->assertContains('close-horizontal', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_SPLIT_OUT_HORIZONTAL':
                    $this->assertContains('open-horizontal', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_SPLIT_IN_VERTICAL':
                    $this->assertContains('close-vertical', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_SPLIT_OUT_VERTICAL':
                    $this->assertContains('open-vertical', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_STRIPS_LEFT_DOWN':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_STRIPS_LEFT_UP':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_STRIPS_RIGHT_DOWN':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_STRIPS_RIGHT_UP':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_WEDGE':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_WIPE_DOWN':
                    $this->assertContains('fade-from-bottom', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_WIPE_LEFT':
                    $this->assertContains('fade-from-left', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_WIPE_RIGHT':
                    $this->assertContains('fade-from-right', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_WIPE_UP':
                    $this->assertContains('fade-from-top', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_ZOOM_IN':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
                case 'TRANSITION_ZOOM_OUT':
                    $this->assertContains('none', $pres->getElementAttribute($element, 'presentation:transition-style', 'content.xml'));
                    break;
            }
        }

        $oTransition->setManualTrigger(true);
        $pres = TestHelperDOCX::getDocument($phpPresentation, 'ODPresentation');
        $this->assertContains('manual', $pres->getElementAttribute($element, 'presentation:transition-type', 'content.xml'));
    }
}
