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
use PhpOffice\PhpPowerpoint\Writer\ODPresentation;
use PhpOffice\PhpPowerpoint\Tests\TestHelperDOCX;
use PhpOffice\PhpPowerpoint\Shape\RichText\Run;

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

    public function testList()
    {
        
    }
}
