<?php

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class PowerPoint2007Test extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $objectPrefix = 'PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\';

        $object = new PowerPoint2007($this->oPresentation);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
        $this->assertInstanceOf("{$objectPrefix}LayoutPack\\PackDefault", $object->getLayoutPack());
    }

    /**
     * Test save
     */
    public function testSave()
    {
        $filename = tempnam(sys_get_temp_dir(), 'PhpPresentation');

        $object = new PowerPoint2007($this->oPresentation);
        $object->save($filename);

        $this->assertTrue(file_exists($filename));

        unlink($filename);
    }

    /**
     * Test save with empty filename
     *
     * @expectedException \Exception
     * @expectedExceptionMessage Filename is empty
     */
    public function testSaveEmptyException()
    {
        $object = new PowerPoint2007($this->oPresentation);
        $object->save('');
    }

    /**
     * Test save with empty assignation
     *
     * @expectedException \Exception
     * @expectedExceptionMessage No PhpPresentation assigned.
     */
    public function testSaveUnassignedException()
    {
        $object = new PowerPoint2007();
        $object->save('filename.pptx');
    }

    /**
     * Test get PhpPresentation exception
     *
     * @expectedException \Exception
     * @expectedExceptionMessage No PhpPresentation assigned.
     */
    public function testGetPhpPresentationException()
    {
        $object = new PowerPoint2007();
        $object->getPhpPresentation();
    }

    /**
     * Test disk caching
     */
    public function testDiskCaching()
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
     * Test set/get disk caching exception
     *
     * @expectedException \Exception
     * @expectedExceptionMessage Directory does not exist: foo
     */
    public function testCachingException()
    {
        $object = new PowerPoint2007($this->oPresentation);
        $object->setUseDiskCaching(true, 'foo');
    }

    /**
     * Test LayoutPack
     * @deprecated 0.7
     */
    public function testLayoutPack()
    {
        $oLayoutPack = $this->getMockBuilder('PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\LayoutPack\\AbstractLayoutPack')->getMock();

        $object = new PowerPoint2007();

        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\LayoutPack\\AbstractLayoutPack", $object->getLayoutPack());
        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007", $object->setLayoutPack());
        $this->assertNull($object->getLayoutPack());
        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007", $object->setLayoutPack($oLayoutPack));
        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\LayoutPack\\AbstractLayoutPack", $object->getLayoutPack());
    }

    public function testZoom()
    {
        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'n', 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'd', 100);
        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'n', 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'd', 100);

        $value = rand(1, 100);
        $this->oPresentation->getPresentationProperties()->setZoom($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'n', $value * 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'd', 100);
        $this->assertZipXmlElementExists('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy');
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'n', $value * 100);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'd', 100);
    }

    public function testFeatureThumbnail()
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

        $xPathManifest = '/Relationships/Relationship[@Target=\'docProps/thumbnail.jpeg\'][@Type=\'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\']';

        $this->assertZipFileExists('_rels/.rels');
        $this->assertZipXmlElementNotExists('_rels/.rels', $xPathManifest);

        $this->oPresentation->getPresentationProperties()->setThumbnailPath($imagePath);
        $this->resetPresentationFile();

        $this->assertZipFileExists('_rels/.rels');
        $this->assertZipXmlElementExists('_rels/.rels', $xPathManifest);
    }
}
