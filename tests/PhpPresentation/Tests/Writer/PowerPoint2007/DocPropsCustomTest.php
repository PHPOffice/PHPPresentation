<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class DocPropsCustomTest extends PhpPresentationTestCase
{
    public function testRender()
    {
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'docProps/custom.xml');
        $this->assertZipXmlElementNotExists($this->oPresentation, 'PowerPoint2007', 'docProps/custom.xml', '/Properties/property[@name="_MarkAsFinal"]');
    }

    public function testMarkAsFinalTrue()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(true);

        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]/vt:bool');
    }

    public function testMarkAsFinalFalse()
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(false);

        $this->assertZipXmlElementNotExists($this->oPresentation, 'PowerPoint2007', 'docProps/custom.xml', '/Properties/property[@name="_MarkAsFinal"]');
    }
}
