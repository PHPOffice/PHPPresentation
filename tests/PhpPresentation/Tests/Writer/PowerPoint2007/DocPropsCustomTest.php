<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class DocPropsCustomTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender()
    {
        $this->assertZipFileExists('docProps/custom.xml');
        $this->assertZipXmlElementNotExists('docProps/custom.xml', '/Properties/property[@name="_MarkAsFinal"]');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalTrue()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(true);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]/vt:bool');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalFalse()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(false);

        $this->assertZipXmlElementNotExists('docProps/custom.xml', '/Properties/property[@name="_MarkAsFinal"]');
        $this->assertIsSchemaECMA376Valid();
    }
}
