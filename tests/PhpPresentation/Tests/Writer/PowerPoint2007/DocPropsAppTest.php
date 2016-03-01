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

class DocPropsAppTest extends \PHPUnit_Framework_TestCase
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
        $this->assertTrue($oXMLDoc->fileExists('docProps/app.xml'));
    }

    public function testCompany()
    {
        $expected = 'aAbBcDeE';

        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getDocumentProperties()->setCompany($expected);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('docProps/app.xml'));
        $this->assertTrue($oXMLDoc->elementExists('/Properties/Company', 'docProps/app.xml'));
        $this->assertEquals($expected, $oXMLDoc->getElement('/Properties/Company', 'docProps/app.xml')->nodeValue);
    }
}
