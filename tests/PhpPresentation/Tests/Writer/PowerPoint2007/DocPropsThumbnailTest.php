<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

/**
 * Class DocPropsThumbnailTest
 * @package PhpPresentation\Tests\Writer\PowerPoint2007
 */
class DocPropsThumbnailTest extends PhpPresentationTestCase
{
    public function testRender()
    {
        $this->assertZipFileNotExists($this->oPresentation, 'PowerPoint2007', 'docProps/thumbnail.jpeg');
    }

    public function testFeatureThumbnail()
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

        $this->oPresentation->getPresentationProperties()->setThumbnailPath($imagePath);
        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'docProps/thumbnail.jpeg');
    }
}
