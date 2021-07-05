<?php

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * Test class for PowerPoint2007.
 *
 * @coversDefaultClass \PowerPoint2007
 */
class PowerPoint2007Test extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new PowerPoint2007($this->oPresentation);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
    }

    /**
     * Test save.
     */
    public function testSave(): void
    {
        $filename = tempnam(sys_get_temp_dir(), 'PhpPresentation');

        $object = new PowerPoint2007($this->oPresentation);
        $object->save($filename);

        $this->assertTrue(file_exists($filename));

        unlink($filename);
    }

    /**
     * Test save with empty filename.
     */
    public function testSaveEmptyException(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('Filename is empty');

        $object = new PowerPoint2007($this->oPresentation);
        $object->save('');
    }

    /**
     * Test save with empty assignation.
     */
    public function testSaveUnassignedException(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('No PhpPresentation assigned.');

        $object = new PowerPoint2007();
        $object->save('filename.pptx');
    }

    /**
     * Test get PhpPresentation exception.
     */
    public function testGetPhpPresentationException(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('No PhpPresentation assigned.');

        $object = new PowerPoint2007();
        $object->getPhpPresentation();
    }

    /**
     * Test disk caching.
     */
    public function testDiskCaching(): void
    {
        $object = new PowerPoint2007($this->oPresentation);
        $this->assertFalse($object->hasDiskCaching());

        $object->setUseDiskCaching(true);
        $this->assertTrue($object->hasDiskCaching());
        $this->assertEquals('./', $object->getDiskCachingDirectory());

        $object->setUseDiskCaching(true, sys_get_temp_dir());
        $this->assertTrue($object->hasDiskCaching());
        $this->assertEquals(sys_get_temp_dir(), $object->getDiskCachingDirectory());
    }

    /**
     * Test set/get disk caching exception.
     */
    public function testCachingException(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('Directory does not exist: foo');

        $object = new PowerPoint2007($this->oPresentation);
        $object->setUseDiskCaching(true, 'foo');
    }

    public function testZoom(): void
    {
        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'n', 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'd', 100);
        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'n', 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'd', 100);
        $this->assertIsSchemaECMA376Valid();

        $value = mt_rand(1, 100);
        $this->oPresentation->getPresentationProperties()->setZoom($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'n', $value * 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'd', 100);
        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'n', $value * 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'd', 100);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testFeatureThumbnail(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        $xPathManifest = '/Relationships/Relationship[@Target=\'docProps/thumbnail.jpeg\'][@Type=\'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\']';

        $this->assertZipFileExists('_rels/.rels');
        $this->assertZipXmlElementNotExists('_rels/.rels', $xPathManifest);
        $this->assertIsSchemaECMA376Valid();

        $this->oPresentation->getPresentationProperties()->setThumbnailPath($imagePath);
        $this->resetPresentationFile();

        $this->assertZipFileExists('_rels/.rels');
        $this->assertZipXmlElementExists('_rels/.rels', $xPathManifest);
        $this->assertIsSchemaECMA376Valid();
    }
}
