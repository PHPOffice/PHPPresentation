<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptPresPropsTest extends PhpPresentationTestCase
{
    public function testRender()
    {
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml');
        $element = '/p:presentationPr/p:extLst/p:ext';
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml', $element);
        $this->assertZipXmlAttributeEquals($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml', $element, 'uri', '{E76CE94A-603C-4142-B9EB-6D1370010A27}');
    }

    public function testLoopContinuously()
    {
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml');
        $element = '/p:presentationPr/p:showPr';
        $this->assertZipXmlElementNotExists($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml', $element);

        $this->oPresentation->getPresentationProperties()->setLoopContinuouslyUntilEsc(true);
        $this->resetPresentationFile();

        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml');
        $element = '/p:presentationPr/p:showPr';
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml', $element);
        $this->assertZipXmlAttributeExists($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml', $element, 'loop');
        $this->assertZipXmlAttributeEquals($this->oPresentation, 'PowerPoint2007', 'ppt/presProps.xml', $element, 'loop', 1);
    }
}
