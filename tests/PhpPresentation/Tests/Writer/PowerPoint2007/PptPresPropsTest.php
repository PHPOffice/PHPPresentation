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

class PptPresPropsTest extends \PHPUnit_Framework_TestCase
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
        $this->assertTrue($oXMLDoc->fileExists('ppt/presProps.xml'));
        $element = '/p:presentationPr/p:extLst/p:ext';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/presProps.xml'));
        $this->assertEquals('{E76CE94A-603C-4142-B9EB-6D1370010A27}', $oXMLDoc->getElementAttribute($element, 'uri', 'ppt/presProps.xml'));
    }

    public function testLoopContinuously()
    {
        $oPhpPresentation = new PhpPresentation();

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('ppt/presProps.xml'));
        $element = '/p:presentationPr/p:showPr';
        $this->assertFalse($oXMLDoc->elementExists($element, 'ppt/presProps.xml'));

        $oPhpPresentation->getPresentationProperties()->setLoopContinuouslyUntilEsc(true);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('ppt/presProps.xml'));
        $element = '/p:presentationPr/p:showPr';
        $this->assertTrue($oXMLDoc->elementExists($element, 'ppt/presProps.xml'));
        $this->assertTrue($oXMLDoc->attributeElementExists($element, 'loop', 'ppt/presProps.xml'));
        $this->assertEquals(1, $oXMLDoc->getElementAttribute($element, 'loop', 'ppt/presProps.xml'));
    }
}
