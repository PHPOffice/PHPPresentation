<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptPresPropsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender()
    {
        $this->assertZipFileExists('ppt/presProps.xml');
        $element = '/p:presentationPr/p:extLst/p:ext';
        $this->assertZipXmlElementExists('ppt/presProps.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/presProps.xml', $element, 'uri', '{E76CE94A-603C-4142-B9EB-6D1370010A27}');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testLoopContinuously()
    {
        $this->assertZipFileExists('ppt/presProps.xml');
        $element = '/p:presentationPr/p:showPr';
        $this->assertZipXmlElementNotExists('ppt/presProps.xml', $element);
        $this->assertIsSchemaECMA376Valid();

        $this->oPresentation->getPresentationProperties()->setLoopContinuouslyUntilEsc(true);
        $this->resetPresentationFile();

        $this->assertZipFileExists('ppt/presProps.xml');
        $element = '/p:presentationPr/p:showPr';
        $this->assertZipXmlElementExists('ppt/presProps.xml', $element);
        $this->assertZipXmlAttributeExists('ppt/presProps.xml', $element, 'loop');
        $this->assertZipXmlAttributeEquals('ppt/presProps.xml', $element, 'loop', 1);
        $this->assertIsSchemaECMA376Valid();
    }
}
