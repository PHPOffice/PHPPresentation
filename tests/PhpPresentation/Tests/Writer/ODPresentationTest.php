<?php

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\ODPresentation;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation
 */
class ODPresentationTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        ;
        $this->oPresentation->getActiveSlide()->createDrawingShape();
        $object = new ODPresentation($this->oPresentation);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\HashTable', $object->getDrawingHashTable());
    }

    /**
     * Test save
     */
    public function testSave()
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
     * Test get PhpPresentation exception
     *
     * @expectedException \Exception
     * @expectedExceptionMessage Filename is empty
     */
    public function testSaveEmpty()
    {
        $object = new ODPresentation();
        $object->save('');
    }

    /**
     * Test get PhpPresentation exception
     *
     * @expectedException \Exception
     * @expectedExceptionMessage No PhpPresentation assigned.
     */
    public function testGetPhpPresentationException()
    {
        $object = new ODPresentation();
        $object->getPhpPresentation();
    }

    /**
     * Test set/get disk caching
     */
    public function testSetGetUseDiskCaching()
    {
        $object = new ODPresentation($this->oPresentation);
        $this->assertFalse($object->hasDiskCaching());

        $object->setUseDiskCaching(true, sys_get_temp_dir());
        $this->assertTrue($object->hasDiskCaching());
        $this->assertEquals(sys_get_temp_dir(), $object->getDiskCachingDirectory());
    }

    /**
     * Test set/get disk caching exception
     *
     * @expectedException \Exception
     */
    public function testSetUseDiskCachingException()
    {
        $object = new ODPresentation($this->oPresentation);
        $object->setUseDiskCaching(true, 'foo');
    }

    public function testFeatureThumbnail()
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

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
