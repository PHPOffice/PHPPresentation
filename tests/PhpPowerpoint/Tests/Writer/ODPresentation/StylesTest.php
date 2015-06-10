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

use PhpOffice\PhpPowerpoint\DocumentLayout;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation;
use PhpOffice\PhpPowerpoint\Tests\TestHelperDOCX;

/**
 * Test class for PhpOffice\PhpPowerpoint\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Writer\ODPresentation
 */
class StylesTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }
    
    public function testDocumentLayout()
    {
        $element = "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties";
    
        $phpPowerPoint = new PhpPowerpoint();
        $oDocumentLayout = new DocumentLayout();
         
        $oDocumentLayout->setDocumentLayout(DocumentLayout::LAYOUT_A4, true);
        $phpPowerPoint->setLayout($oDocumentLayout);
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $this->assertTrue($pres->elementExists($element, 'styles.xml'));
        $this->assertEquals('landscape', $pres->getElementAttribute($element, 'style:print-orientation', 'styles.xml'));
        
        $oDocumentLayout->setDocumentLayout(DocumentLayout::LAYOUT_A4, false);
        $phpPowerPoint->setLayout($oDocumentLayout);
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $this->assertTrue($pres->elementExists($element, 'styles.xml'));
        $this->assertEquals('portrait', $pres->getElementAttribute($element, 'style:print-orientation', 'styles.xml'));
    }
    
    public function testCustomDocumentLayout()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oDocumentLayout = new DocumentLayout();
        $oDocumentLayout->setDocumentLayout(array('cx' => rand(1, 100),'cy' => rand(1, 100),));
        $phpPowerPoint->setLayout($oDocumentLayout);
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        
        $element = "/office:document-styles/office:automatic-styles/style:page-layout";
        $this->assertTrue($pres->elementExists($element, 'styles.xml'));
        $this->assertEquals('sPL0', $pres->getElementAttribute($element, 'style:name', 'styles.xml'));
        
        $element = "/office:document-styles/office:master-styles/style:master-page";
        $this->assertTrue($pres->elementExists($element, 'styles.xml'));
        $this->assertEquals('sPL0', $pres->getElementAttribute($element, 'style:page-layout-name', 'styles.xml'));
    }
    
    public function testFillGradientLinearRichText()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FFFF7700'))->setEndColor(new Color('FFFFFFFF'));
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-styles/office:styles/draw:gradient';
        $this->assertEquals('gradient_'.$oShape->getFill()->getHashCode(), $pres->getElementAttribute($element, 'draw:name', 'styles.xml'));
        
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('gradient', $pres->getElementAttribute($element, 'draw:fill', 'content.xml'));
        $this->assertEquals('gradient_'.$oShape->getFill()->getHashCode(), $pres->getElementAttribute($element, 'draw:fill-gradient-name', 'content.xml'));
    }
    
    public function testFillSolidRichText()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF4672A8'));
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'gr1\']/style:graphic-properties';
        $this->assertTrue($pres->elementExists($element, 'content.xml'));
        $this->assertEquals('solid', $pres->getElementAttribute($element, 'draw:fill', 'content.xml'));
        $this->assertEquals('#'.$oShape->getFill()->getStartColor()->getRGB(), $pres->getElementAttribute($element, 'draw:fill-color', 'content.xml'));
        $this->assertEquals('#'.$oShape->getFill()->getEndColor()->getRGB(), $pres->getElementAttribute($element, 'draw:fill-color', 'content.xml'));
    }
    
    public function testGradientTable()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FFFF7700'))->setEndColor(new Color('FFFFFFFF'));
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = "/office:document-styles/office:styles/draw:gradient";
        $this->assertEquals('gradient_'.$oCell->getFill()->getHashCode(), $pres->getElementAttribute($element, 'draw:name', 'styles.xml'));
    }
    
    public function testStrokeDash()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setLineStyle(Border::LINE_SINGLE);
        $arrayDashStyle = array(
            Border::DASH_DASH,
            Border::DASH_DASHDOT,
            Border::DASH_DOT,
            Border::DASH_LARGEDASH,
            Border::DASH_LARGEDASHDOT,
            Border::DASH_LARGEDASHDOTDOT,
            Border::DASH_SYSDASH,
            Border::DASH_SYSDASHDOT,
            Border::DASH_SYSDASHDOTDOT,
            Border::DASH_SYSDOT,
        );
        
        foreach ($arrayDashStyle as $style){
            $oRichText1->getBorder()->setDashStyle($style);

            $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
            $element = '/office:document-styles/office:styles/draw:stroke-dash[@draw:name=\'strokeDash_'.$style.'\']';
            $this->assertTrue($pres->elementExists($element, 'styles.xml'));
            $this->assertEquals('rect', $pres->getElementAttribute($element, 'draw:style', 'styles.xml'));
            $this->assertTrue($pres->attributeElementExists($element, 'draw:distance', 'styles.xml'));
            
            switch($style){
                case Border::DASH_DOT:
                case Border::DASH_SYSDOT:
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots1', 'styles.xml'));
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots1-length', 'styles.xml'));
                    break;
                case Border::DASH_DASH:
                case Border::DASH_LARGEDASH:
                case Border::DASH_SYSDASH:
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots2', 'styles.xml'));
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots2-length', 'styles.xml'));
                    break;
                case Border::DASH_DASHDOT:
                case Border::DASH_LARGEDASHDOT:
                case Border::DASH_LARGEDASHDOTDOT:
                case Border::DASH_SYSDASHDOT:
                case Border::DASH_SYSDASHDOTDOT:
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots1', 'styles.xml'));
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots1-length', 'styles.xml'));
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots2', 'styles.xml'));
                    $this->assertTrue($pres->attributeElementExists($element, 'draw:dots2-length', 'styles.xml'));
                    break;
            }
        }
    }
}
