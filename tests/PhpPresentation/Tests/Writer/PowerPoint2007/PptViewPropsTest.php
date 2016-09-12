<?php
/**
 * Created by PhpStorm.
 * User: lefevre_f
 * Date: 01/03/2016
 * Time: 12:35
 */

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

class PptViewPropsTest extends \PHPUnit_Framework_TestCase
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
        $expectedElement = '/p:viewPr';

        $oPhpPresentation = new PhpPresentation();

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('ppt/viewProps.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'ppt/viewProps.xml'));
        $this->assertEquals('0', $oXMLDoc->getElementAttribute($expectedElement, 'showComments', 'ppt/viewProps.xml'));
        $this->assertEquals(PresentationProperties::VIEW_SLIDE, $oXMLDoc->getElementAttribute($expectedElement, 'lastView', 'ppt/viewProps.xml'));
    }

    public function testCommentVisible()
    {
        $expectedElement ='/p:viewPr';

        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getPresentationProperties()->setCommentVisible(true);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('ppt/viewProps.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'ppt/viewProps.xml'));
        $this->assertEquals(1, $oXMLDoc->getElementAttribute($expectedElement, 'showComments', 'ppt/viewProps.xml'));
    }

    public function testLastView()
    {
        $expectedElement ='/p:viewPr';
        $expectedLastView = PresentationProperties::VIEW_OUTLINE;

        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getPresentationProperties()->setLastView($expectedLastView);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('ppt/viewProps.xml'));
        $this->assertTrue($oXMLDoc->elementExists($expectedElement, 'ppt/viewProps.xml'));
        $this->assertEquals($expectedLastView, $oXMLDoc->getElementAttribute($expectedElement, 'lastView', 'ppt/viewProps.xml'));
    }
}
