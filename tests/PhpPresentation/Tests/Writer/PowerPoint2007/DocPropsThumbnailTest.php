<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

/**
 * Class DocPropsThumbnailTest.
 */
class DocPropsThumbnailTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender(): void
    {
        $this->assertZipFileNotExists('docProps/thumbnail.jpeg');
        $this->assertZipXmlElementNotExists('_rels/.rels', '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"]');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testFeatureThumbnailFile(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        $this->oPresentation->getPresentationProperties()
            ->setThumbnailPath($imagePath, PresentationProperties::THUMBNAIL_FILE);
        $this->assertZipFileExists('docProps/thumbnail.jpeg');
        $this->assertZipXmlElementExists('_rels/.rels', '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"]');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testFeatureThumbnailFileNotExisting(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'NotExistingFile.png';

        $this->oPresentation->getPresentationProperties()
            ->setThumbnailPath($imagePath, PresentationProperties::THUMBNAIL_FILE);
        $this->assertZipFileNotExists('docProps/thumbnail.jpeg');
        $this->assertZipXmlElementNotExists('_rels/.rels', '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"]');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testFeatureThumbnailData(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        $this->oPresentation->getPresentationProperties()
            ->setThumbnailPath('', PresentationProperties::THUMBNAIL_DATA, file_get_contents($imagePath));
        $this->assertZipFileExists('docProps/thumbnail.jpeg');
        $this->assertZipXmlElementExists('_rels/.rels', '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"]');
        $this->assertIsSchemaECMA376Valid();
    }
}
