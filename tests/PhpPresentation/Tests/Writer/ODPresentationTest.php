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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\ODPresentation;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\ODPresentation
 */
class ODPresentationTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $this->oPresentation->getActiveSlide()->createDrawingShape();
        $object = new ODPresentation($this->oPresentation);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\HashTable', $object->getDrawingHashTable());
    }

    /**
     * Test save.
     */
    public function testSave(): void
    {
        $filename = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        $imageFile = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png';

        $slide = $this->oPresentation->getActiveSlide();
        $slide->createRichTextShape();
        $slide->createLineShape(10, 10, 10, 10);
        $slide->createChartShape()->getPlotArea()->setType(new Bar3D());
        $slide->createDrawingShape()->setName('Drawing')->setPath($imageFile);
        $slide->createTableShape()->createRow();

        $object = new ODPresentation($this->oPresentation);
        $object->save($filename);

        $this->assertTrue(file_exists($filename));

        unlink($filename);
    }

    /**
     * Test get PhpPresentation exception.
     */
    public function testSaveEmpty(): void
    {
        $this->expectException(InvalidParameterException::class);
        $this->expectExceptionMessage('The parameter pFilename can\'t have the value ""');

        $object = new ODPresentation();
        $object->save('');
    }

    /**
     * Test set/get disk caching.
     */
    public function testSetGetUseDiskCaching(): void
    {
        $object = new ODPresentation($this->oPresentation);
        $this->assertFalse($object->hasDiskCaching());

        $object->setUseDiskCaching(true, sys_get_temp_dir());
        $this->assertTrue($object->hasDiskCaching());
        $this->assertEquals(sys_get_temp_dir(), $object->getDiskCachingDirectory());
    }

    /**
     * Test set/get disk caching exception.
     */
    public function testSetUseDiskCachingException(): void
    {
        $this->expectException(DirectoryNotFoundException::class);

        $object = new ODPresentation($this->oPresentation);
        $object->setUseDiskCaching(true, 'foo');
    }

    public function testFeatureThumbnail(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        $xPathManifest = '/manifest:manifest/manifest:file-entry[@manifest:media-type=\'image/png\'][@manifest:full-path=\'Thumbnails/thumbnail.png\']';

        $this->assertZipFileNotExists('Thumbnails/thumbnail.png');
        $this->assertZipFileExists('META-INF/manifest.xml');
        $this->assertZipXmlElementNotExists('META-INF/manifest.xml', $xPathManifest);

        $this->oPresentation->getPresentationProperties()->setThumbnailPath($imagePath);
        $this->resetPresentationFile();

        $this->assertZipFileExists('Thumbnails/thumbnail.png');
        $this->assertZipFileExists('META-INF/manifest.xml');
        $this->assertZipXmlElementExists('META-INF/manifest.xml', $xPathManifest);
    }
}
