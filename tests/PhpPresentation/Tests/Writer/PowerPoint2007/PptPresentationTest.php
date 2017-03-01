<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptPresentationTest extends PhpPresentationTestCase
{
    public function testRender()
    {
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/presentation.xml');
    }
}
