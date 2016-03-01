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

class DocPropsCoreTest extends \PHPUnit_Framework_TestCase
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
        $this->assertTrue($oXMLDoc->fileExists('docProps/core.xml'));
    }

    public function testDocumentProperties()
    {
        $expected = 'aAbBcDeE';

        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getDocumentProperties()->setCreator($expected);
        $oPhpPresentation->getDocumentProperties()->setTitle($expected);
        $oPhpPresentation->getDocumentProperties()->setDescription($expected);
        $oPhpPresentation->getDocumentProperties()->setSubject($expected);
        $oPhpPresentation->getDocumentProperties()->setKeywords($expected);
        $oPhpPresentation->getDocumentProperties()->setCategory($expected);

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('docProps/core.xml'));

        $this->assertTrue($oXMLDoc->elementExists('/cp:coreProperties/dc:creator', 'docProps/core.xml'));
        $this->assertEquals($expected, $oXMLDoc->getElement('/cp:coreProperties/dc:creator', 'docProps/core.xml')->nodeValue);

        $this->assertTrue($oXMLDoc->elementExists('/cp:coreProperties/dc:title', 'docProps/core.xml'));
        $this->assertEquals($expected, $oXMLDoc->getElement('/cp:coreProperties/dc:title', 'docProps/core.xml')->nodeValue);

        $this->assertTrue($oXMLDoc->elementExists('/cp:coreProperties/dc:description', 'docProps/core.xml'));
        $this->assertEquals($expected, $oXMLDoc->getElement('/cp:coreProperties/dc:description', 'docProps/core.xml')->nodeValue);

        $this->assertTrue($oXMLDoc->elementExists('/cp:coreProperties/dc:subject', 'docProps/core.xml'));
        $this->assertEquals($expected, $oXMLDoc->getElement('/cp:coreProperties/dc:subject', 'docProps/core.xml')->nodeValue);

        $this->assertTrue($oXMLDoc->elementExists('/cp:coreProperties/cp:keywords', 'docProps/core.xml'));
        $this->assertEquals($expected, $oXMLDoc->getElement('/cp:coreProperties/cp:keywords', 'docProps/core.xml')->nodeValue);

        $this->assertTrue($oXMLDoc->elementExists('/cp:coreProperties/cp:category', 'docProps/core.xml'));
        $this->assertEquals($expected, $oXMLDoc->getElement('/cp:coreProperties/cp:category', 'docProps/core.xml')->nodeValue);
    }

    public function testMarkAsFinal()
    {
        $oPhpPresentation = new PhpPresentation();

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->elementExists('/cp:coreProperties/cp:contentStatus', 'docProps/core.xml'));

        $oPhpPresentation->getPresentationProperties()->markAsFinal(true);

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->elementExists('/cp:coreProperties/cp:contentStatus', 'docProps/core.xml'));
        $this->assertEquals('Final', $pres->getElement('/cp:coreProperties/cp:contentStatus', 'docProps/core.xml')->nodeValue);

        $oPhpPresentation->getPresentationProperties()->markAsFinal(false);

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->elementExists('/cp:coreProperties/cp:contentStatus', 'docProps/core.xml'));
    }
}
