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

namespace PhpOffice\PhpPowerpoint\Tests\Writer\ODPresentation;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\RichText\Run;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Bullet;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation;
use PhpOffice\PhpPowerpoint\Tests\TestHelperDOCX;

/**
 * Test class for PhpOffice\PhpPowerpoint\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Writer\ODPresentation\Manifest
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

    public function testList()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        
        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testInnerList()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        
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
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        
        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:list/text:list-item/text:p/text:span';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testParagraphRichText()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('Alpha');
        $oRichText->createBreak();
        $oRichText->createText('Beta');
        $oRichText->createBreak();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:line-break';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:a';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('http://www.google.fr', $pres->getElementAttribute($element, 'xlink:href', 'content.xml'));
    }

    public function testListWithRichText()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRun = $oRichText->createTextRun('Alpha');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
        $oRichText->createBreak();
        $oRichText->createTextRun('Beta');
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:a';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:line-break';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }

    public function testRichtextAutoShrink()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertFalse($pres->attributeElementExists($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertFalse($pres->attributeElementExists($element, 'draw:auto-grow-width', 'content.xml'));

        $oRichText1->setAutoShrinkHorizontal(false);
        $oRichText1->setAutoShrinkVertical(true);
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-width', 'content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'draw:auto-grow-width', 'content.xml'));
        

        $oRichText1->setAutoShrinkHorizontal(true);
        $oRichText1->setAutoShrinkVertical(false);
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertTrue($pres->attributeElementExists($element, 'draw:auto-grow-width', 'content.xml'));
        $this->assertEquals('false', $pres->getElementAttribute($element, 'draw:auto-grow-height', 'content.xml'));
        $this->assertEquals('true', $pres->getElementAttribute($element, 'draw:auto-grow-width', 'content.xml'));
    }
    
    public function testStyleAlignment()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
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
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        
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
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setBold(true);
        
        $expectedHashCode = $oRun->getFont()->getHashCode();
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'T_'.$expectedHashCode.'\']/style:text-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('bold', $pres->getElementAttribute($element, 'fo:font-weight', 'content.xml'));
    }
    
    public function testTable()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oSlide->createTableShape();
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
    }
    
    public function testTableWithColspan()
    {
        $value = rand(2, 100);
        
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape($value);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->setColSpan($value);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals($value, $pres->getElementAttribute($element, 'table:number-columns-spanned', 'content.xml'));
    }
    
    public function testTableWithText()
    {
        $oRun = new Run();
        $oRun->setText('Test');
        
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->addText($oRun);
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('Test', $pres->getElement($element, 'content.xml')->nodeValue);
    }
}
