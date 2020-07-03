<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptViewPropsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender()
    {
        $expectedElement = '/p:viewPr';

        $this->assertZipFileExists('ppt/viewProps.xml');
        $this->assertZipXmlElementExists('ppt/viewProps.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'showComments', 0);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'lastView', PresentationProperties::VIEW_SLIDE);
    }

    public function testCommentVisible()
    {
        $expectedElement ='/p:viewPr';

        $this->oPresentation->getPresentationProperties()->setCommentVisible(true);

        $this->assertZipFileExists('ppt/viewProps.xml');
        $this->assertZipXmlElementExists('ppt/viewProps.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'showComments', 1);
    }

    public function testLastView()
    {
        $expectedElement ='/p:viewPr';
        $expectedLastView = PresentationProperties::VIEW_OUTLINE;

        $this->oPresentation->getPresentationProperties()->setLastView($expectedLastView);

        $this->assertZipFileExists('ppt/viewProps.xml');
        $this->assertZipXmlElementExists('ppt/viewProps.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'lastView', $expectedLastView);
    }
}
