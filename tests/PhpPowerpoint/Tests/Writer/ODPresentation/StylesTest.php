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
    
    public function testGradient()
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
}
