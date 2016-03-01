<?php
/**
 * Created by PhpStorm.
 * User: lefevre_f
 * Date: 01/03/2016
 * Time: 12:35
 */

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

class PptTablePropsTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }

    public function testRender()
    {
        $oPhpPresentation = new PhpPresentation();

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('ppt/tableStyles.xml'));
        $element = '/a:tblStyleLst';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/tableStyles.xml'));
        $this->assertEquals('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', $oXMLDoc->getElementAttribute($element, 'def', 'ppt/tableStyles.xml'));
    }
}
