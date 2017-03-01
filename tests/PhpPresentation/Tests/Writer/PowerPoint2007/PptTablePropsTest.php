<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptTablePropsTest extends PhpPresentationTestCase
{
    public function testRender()
    {
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/tableStyles.xml');
        $element = '/a:tblStyleLst';
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'ppt/tableStyles.xml', $element);
        $this->assertZipXmlAttributeEquals($this->oPresentation, 'PowerPoint2007', 'ppt/tableStyles.xml', $element, 'def', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
    }
}
