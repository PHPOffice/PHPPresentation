<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptTablePropsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender()
    {
        $this->assertZipFileExists('ppt/tableStyles.xml');
        $element = '/a:tblStyleLst';
        $this->assertZipXmlElementExists('ppt/tableStyles.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/tableStyles.xml', $element, 'def', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
        $this->assertIsSchemaECMA376Valid();
    }
}
