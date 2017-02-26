<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class DocPropsCoreTest extends PhpPresentationTestCase
{
    public function testRender()
    {
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml');
        $this->assertZipXmlElementNotExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
    }

    public function testDocumentProperties()
    {
        $expected = 'aAbBcDeE';

        $this->oPresentation->getDocumentProperties()
            ->setCreator($expected)
            ->setTitle($expected)
            ->setDescription($expected)
            ->setSubject($expected)
            ->setKeywords($expected)
            ->setCategory($expected);

        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:creator');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:creator', $expected);
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:title');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:title', $expected);
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:description');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:description', $expected);
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:subject');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/dc:subject', $expected);
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:keywords');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:keywords', $expected);
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:category');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:category', $expected);
    }

    public function testMarkAsFinalTrue()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(true);

        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:contentStatus', 'Final');
    }

    public function testMarkAsFinalFalse()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(false);

        $this->assertZipXmlElementNotExists($this->oPresentation, 'PowerPoint2007', 'docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
    }
}
