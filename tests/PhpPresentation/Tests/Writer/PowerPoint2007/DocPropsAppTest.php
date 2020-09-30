<?php
/**
 * Created by PhpStorm.
 * User: lefevre_f
 * Date: 01/03/2016
 * Time: 12:35
 */

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class DocPropsAppTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender()
    {
        $this->assertZipFileExists('docProps/app.xml');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testCompany()
    {
        $expected = 'aAbBcDeE';

        $this->oPresentation->getDocumentProperties()->setCompany($expected);

        $this->assertZipFileExists('docProps/app.xml');
        $this->assertZipXmlElementExists('docProps/app.xml', '/Properties/Company');
        $this->assertZipXmlElementEquals('docProps/app.xml', '/Properties/Company', $expected);
        $this->assertIsSchemaECMA376Valid();
    }
}
