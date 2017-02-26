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
    public function testRender()
    {
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'docProps/app.xml');
    }

    public function testCompany()
    {
        $expected = 'aAbBcDeE';

        $this->oPresentation->getDocumentProperties()->setCompany($expected);

        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'docProps/app.xml');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'docProps/app.xml', '/Properties/Company');
        $this->assertZipXmlElementEquals($this->oPresentation, 'PowerPoint2007', 'docProps/app.xml', '/Properties/Company', $expected);
    }
}
