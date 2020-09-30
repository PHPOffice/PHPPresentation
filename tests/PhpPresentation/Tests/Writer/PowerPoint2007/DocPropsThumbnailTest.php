<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

/**
 * Class DocPropsThumbnailTest
 * @package PhpPresentation\Tests\Writer\PowerPoint2007
 */
class DocPropsThumbnailTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender()
    {
        $this->assertZipFileNotExists('docProps/thumbnail.jpeg');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testFeatureThumbnail()
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

        $this->oPresentation->getPresentationProperties()->setThumbnailPath($imagePath);
        $this->assertZipFileExists('docProps/thumbnail.jpeg');
        $this->assertIsSchemaECMA376Valid();
    }
}
