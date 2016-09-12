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

class DocPropsCustomTest extends \PHPUnit_Framework_TestCase
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
        $this->assertTrue($oXMLDoc->fileExists('docProps/custom.xml'));
    }

    public function testMarkAsFinal()
    {
        $oPhpPresentation = new PhpPresentation();

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->elementExists('/Properties/property[@name="_MarkAsFinal"]', 'docProps/custom.xml'));

        $oPhpPresentation->getPresentationProperties()->markAsFinal(true);

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->elementExists('/Properties', 'docProps/custom.xml'));
        $this->assertTrue($pres->elementExists('/Properties/property', 'docProps/custom.xml'));
        $this->assertTrue($pres->elementExists('/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]', 'docProps/custom.xml'));
        $this->assertTrue($pres->elementExists('/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]/vt:bool', 'docProps/custom.xml'));

        $oPhpPresentation->getPresentationProperties()->markAsFinal(false);

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->elementExists('/Properties/property[@name="_MarkAsFinal"]', 'docProps/custom.xml'));
    }
}
