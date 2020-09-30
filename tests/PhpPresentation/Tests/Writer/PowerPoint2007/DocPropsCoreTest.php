<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class DocPropsCoreTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender()
    {
        $this->assertZipFileExists('docProps/core.xml');
        $this->assertZipXmlElementNotExists('docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
        $this->assertIsSchemaECMA376Valid();
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

        $this->assertZipFileExists('docProps/core.xml');
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:creator');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:creator', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:title');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:title', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:description');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:description', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:subject');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:subject', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/cp:keywords');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/cp:keywords', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/cp:category');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/cp:category', $expected);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalTrue()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(true);

        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/cp:contentStatus', 'Final');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalFalse()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(false);

        $this->assertZipXmlElementNotExists('docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
        $this->assertIsSchemaECMA376Valid();
    }
}
